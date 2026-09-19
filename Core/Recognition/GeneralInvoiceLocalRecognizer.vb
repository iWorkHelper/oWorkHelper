Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.RegularExpressions

''' <summary>
''' 常规发票（非滴滴）专用本地识别器。基于「逻辑行 + 字段候选 + 评分择优」，
''' 优先使用坐标行重建（有坐标时购/销名称按 X 就近归栏），无坐标时按归一化文本换行 + 行序就近。
''' 输出统一的 InvoiceRecognitionResult：成功 / 部分成功 / 需要 OCR / 失败。
''' 商品明细解析失败不影响头部字段；单条明细失败不影响整体。
''' 不改动滴滴发票与行程单的既有识别逻辑（由 LocalTextInvoiceRecognizer 保留处理）。
''' </summary>
Public Class GeneralInvoiceLocalRecognizer

    Private Const RecognizerLabel As String = "GeneralInvoiceLocalRecognizer"

    ''' <summary>
    ''' 名称值的截断停止词：避免把地址/税号/开户行并入名称；并在遇到下一个「名称/销售方/购买方/项目/价税」
    ''' 等字段边界时截断，防止两栏版式（购销并排）时第一个名称吞并到第二个名称。
    ''' </summary>
    Private Const NameStop As String = "(?=统一社会信用代码|纳税人识别号|地址|开户行|开户|电话|账号|销售方|购买方|销方|购方|名称|项目名称|价税|开票日期|$)"

    Public Function Recognize(context As RecognitionContext) As InvoiceRecognitionResult
        Dim result As New InvoiceRecognitionResult()
        result.RecognizerName = RecognizerLabel
        result.Source = RecognitionSource.LocalText
        result.DocumentType = InvoiceDocumentType.VatInvoice

        Try
            Dim rawText As String = If(context Is Nothing, Nothing, context.ExtractedText)
            If String.IsNullOrWhiteSpace(rawText) Then
                Return InvoiceRecognitionResult.NeedsOcr("本地文本层为空，疑似图片型 PDF，需 OCR。", rawText)
            End If
            result.RawText = rawText

            Dim normalized As String = LocalTextNormalizer.Normalize(rawText)
            Dim lines As List(Of PdfTextLine) = BuildLogicalLines(context, normalized)

            Dim parse As GeneralInvoiceParseResult = ParseFields(lines, normalized)

            ' 组装 InvoiceInfo
            Dim inv As InvoiceInfo = result.Invoice
            inv.InvoiceNumber = Adopt(result, parse, InvoiceFieldNames.InvoiceNumber)
            inv.InvoiceCode = Adopt(result, parse, InvoiceFieldNames.InvoiceCode)
            inv.InvoiceDate = Adopt(result, parse, InvoiceFieldNames.InvoiceDate)
            inv.SellerName = Adopt(result, parse, InvoiceFieldNames.SellerName)
            inv.BuyerName = Adopt(result, parse, InvoiceFieldNames.BuyerName)
            inv.SellerTaxId = Adopt(result, parse, InvoiceFieldNames.SellerTaxId)
            inv.BuyerTaxId = Adopt(result, parse, InvoiceFieldNames.BuyerTaxId)
            inv.TotalWithTax = Adopt(result, parse, InvoiceFieldNames.TotalWithTax)
            inv.TaxRate = Adopt(result, parse, InvoiceFieldNames.TaxRate)
            inv.TaxAmount = Adopt(result, parse, InvoiceFieldNames.TaxAmount)
            If Not String.IsNullOrEmpty(parse.InvoiceTypeText) Then inv.SetExtendedField("发票类型", parse.InvoiceTypeText)

            ' 商品明细（失败不影响头部）
            Try
                inv.LineItems = New GeneralInvoiceLineItemParser().Parse(lines)
                If inv.LineItems.Count > 0 AndAlso String.IsNullOrWhiteSpace(inv.ItemName) Then
                    inv.ItemName = inv.LineItems(0).ItemName
                End If
            Catch exItem As Exception
                AppLogger.Warn("常规发票明细解析异常（忽略）：" & exItem.Message)
            End Try

            ' 明细金额合计与价税合计一致性（仅告警）
            CheckAmountConsistency(inv, parse)

            ' 状态判定：命名核心字段 开票日期 + 金额 + 销售方名称
            Dim missing As List(Of String) = KeyFieldEvaluator.GetMissingGeneralInvoiceNamingFields(inv)
            result.MissingKeyFields = missing

            If inv.HasAnyCoreField() Then
                If missing.Count = 0 Then
                    result.Status = RecognitionStatus.Success
                    result.Message = "常规发票本地解析成功（开票日期/金额/销售方齐全）。"
                Else
                    result.Status = RecognitionStatus.PartialSuccess
                    result.Message = "常规发票本地解析部分成功，缺少：" & String.Join("、", missing.ToArray()) & "（建议回退在线 OCR）。"
                End If
            Else
                Return InvoiceRecognitionResult.NeedsOcr("常规发票本地未解析到关键字段，建议 OCR。", normalized)
            End If

            AppLogger.Info(String.Format("常规发票本地识别：类型={0}, 状态={1}, 字段数={2}, 明细={3}, 缺失={4}",
                                         If(parse.InvoiceTypeText, "增值税发票"), result.Status, result.Fields.Count,
                                         inv.LineItems.Count, If(missing.Count = 0, "无", String.Join("/", missing.ToArray()))))
            Return result

        Catch ex As Exception
            AppLogger.Error("常规发票本地识别异常。", ex)
            Return InvoiceRecognitionResult.Failure("常规发票本地识别异常: " & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

    ''' <summary>暴露给 OfflineTester 诊断：返回带全部候选/明细的解析结果。</summary>
    Public Function Diagnose(context As RecognitionContext) As GeneralInvoiceParseResult
        Dim rawText As String = If(context Is Nothing, Nothing, context.ExtractedText)
        If String.IsNullOrWhiteSpace(rawText) Then Return New GeneralInvoiceParseResult()
        Dim normalized As String = LocalTextNormalizer.Normalize(rawText)
        Dim lines As List(Of PdfTextLine) = BuildLogicalLines(context, normalized)
        Dim parse As GeneralInvoiceParseResult = ParseFields(lines, normalized)
        Try
            parse.LineItems = New GeneralInvoiceLineItemParser().Parse(lines)
        Catch
        End Try
        Return parse
    End Function

    ' ========== 逻辑行 ==========

    ''' <summary>
    ''' 构建逻辑行：优先坐标行重建（含 X 坐标）；无则按归一化文本换行切分（合成单词，X=0）。
    ''' </summary>
    Private Function BuildLogicalLines(context As RecognitionContext, normalized As String) As List(Of PdfTextLine)
        Dim lines As New List(Of PdfTextLine)()
        Dim extract As PdfTextExtractResult = If(context Is Nothing, Nothing, context.TextExtractResult)
        If extract IsNot Nothing AndAlso extract.Lines IsNot Nothing AndAlso extract.Lines.Count > 0 Then
            lines.AddRange(extract.Lines)
            Return lines
        End If

        Dim arr As String() = normalized.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ChrW(10))
        Dim y As Double = 10000
        For Each raw As String In arr
            Dim t As String = If(raw, "").Trim()
            If t.Length = 0 Then y -= 12 : Continue For
            Dim ln As New PdfTextLine With {.Y = y, .PageIndex = 1}
            ln.Words.Add(New PdfTextWord With {.Text = t, .X = 0, .Right = 0, .Y = y})
            lines.Add(ln)
            y -= 12
        Next
        Return lines
    End Function

    Private Function Compact(ln As PdfTextLine) As String
        If ln Is Nothing OrElse ln.Text Is Nothing Then Return ""
        Return ln.Text.Replace(" ", "")
    End Function

    ' ========== 字段解析 ==========

    Private Function ParseFields(lines As List(Of PdfTextLine), normalized As String) As GeneralInvoiceParseResult
        Dim parse As New GeneralInvoiceParseResult()
        Dim compactAll As String = If(normalized, "").Replace(" ", "").Replace(vbLf, "").Replace(vbCr, "")

        parse.InvoiceTypeText = DetectInvoiceType(compactAll)

        ParseInvoiceNumber(lines, compactAll, parse)
        ParseInvoiceCode(lines, parse)
        ParseInvoiceDate(lines, compactAll, parse)
        ParseParties(lines, parse)
        ParseTotalWithTax(lines, compactAll, parse)
        ParseTaxRateAndAmount(lines, compactAll, parse)

        parse.Chosen = GeneralInvoiceCandidateScorer.ChooseBest(parse.Candidates)
        Return parse
    End Function

    Private Function DetectInvoiceType(compactAll As String) As String
        Dim c As String = compactAll
        If c.Contains("增值税专用发票") Then Return "增值税专用发票"
        If c.Contains("增值税电子专用发票") Then Return "增值税电子专用发票"
        If c.Contains("增值税电子普通发票") Then Return "增值税电子普通发票"
        If c.Contains("电子发票(普通发票)") OrElse c.Contains("电子发票（普通发票）") Then Return "电子发票（普通发票）"
        If c.Contains("增值税普通发票") Then Return "增值税普通发票"
        If c.Contains("数电") Then Return "数电票"
        If c.Contains("电子发票") Then Return "电子发票"
        If c.Contains("普通发票") Then Return "普通发票"
        Return "增值税发票"
    End Function

    Private Sub ParseInvoiceNumber(lines As List(Of PdfTextLine), compactAll As String, parse As GeneralInvoiceParseResult)
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If c.Contains("发票代码") AndAlso Not c.Contains("发票号码") Then Continue For ' 避免从代码行取号码
            ' 数字后不能紧跟字母（避免取到「91310000MA…」这类税号/信用代码的数字前缀）
            Dim m As Match = Regex.Match(c, "发票号码[:：]?\s*([0-9]{6,20})(?![0-9A-Za-z])")
            If m.Success Then
                AddCand(parse, InvoiceFieldNames.InvoiceNumber, m.Groups(1).Value, lines(i), i, "label-same-line",
                        GeneralInvoiceCandidateScorer.LabelProximityBonus(0) + GeneralInvoiceCandidateScorer.InvoiceNumberPlausibility(m.Groups(1).Value) _
                        + GeneralInvoiceCandidateScorer.LongNumberContextPenalty(c), "标签同行")
            ElseIf c.EndsWith("发票号码", StringComparison.Ordinal) OrElse c = "发票号码" Then
                ' 号码可能在下一行
                If i + 1 < lines.Count Then
                    Dim nx As Match = Regex.Match(Compact(lines(i + 1)), "^([0-9]{8,20})$")
                    If nx.Success Then
                        AddCand(parse, InvoiceFieldNames.InvoiceNumber, nx.Groups(1).Value, lines(i + 1), i + 1, "label-next-line",
                                GeneralInvoiceCandidateScorer.LabelProximityBonus(1) + GeneralInvoiceCandidateScorer.InvoiceNumberPlausibility(nx.Groups(1).Value), "标签下一行")
                    End If
                End If
            End If
        Next
        ' 全局兜底：紧跟「发票号码」的数字（同样排除数字后紧跟字母的情形）
        Dim g As Match = Regex.Match(compactAll, "发票号码[:：]?([0-9]{8,20})(?![0-9A-Za-z])")
        If g.Success Then
            AddCand(parse, InvoiceFieldNames.InvoiceNumber, g.Groups(1).Value, Nothing, -1, "global",
                    2 + GeneralInvoiceCandidateScorer.InvoiceNumberPlausibility(g.Groups(1).Value), "全局兜底")
        End If
    End Sub

    Private Sub ParseInvoiceCode(lines As List(Of PdfTextLine), parse As GeneralInvoiceParseResult)
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            Dim m As Match = Regex.Match(c, "发票代码[:：]?\s*([0-9]{10,12})")
            If m.Success Then
                AddCand(parse, InvoiceFieldNames.InvoiceCode, m.Groups(1).Value, lines(i), i, "label-same-line",
                        GeneralInvoiceCandidateScorer.LabelProximityBonus(0) + GeneralInvoiceCandidateScorer.InvoiceCodePlausibility(m.Groups(1).Value), "标签同行")
            End If
        Next
    End Sub

    Private Sub ParseInvoiceDate(lines As List(Of PdfTextLine), compactAll As String, parse As GeneralInvoiceParseResult)
        Dim datePat As String = "(\d{4}\s*[-/.年]\s*\d{1,2}\s*[-/.月]\s*\d{1,2}\s*日?)"
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If Not c.Contains("开票日期") Then Continue For
            Dim m As Match = Regex.Match(c, "开票日期[:：]?\s*" & datePat)
            If m.Success Then
                AddCand(parse, InvoiceFieldNames.InvoiceDate, m.Groups(1).Value, lines(i), i, "label-same-line",
                        GeneralInvoiceCandidateScorer.LabelProximityBonus(0) + GeneralInvoiceCandidateScorer.DateFormatBonus(m.Groups(1).Value), "标签同行")
            ElseIf i + 1 < lines.Count Then
                Dim nx As Match = Regex.Match(Compact(lines(i + 1)), "^" & datePat)
                If nx.Success Then
                    AddCand(parse, InvoiceFieldNames.InvoiceDate, nx.Groups(1).Value, lines(i + 1), i + 1, "label-next-line",
                            GeneralInvoiceCandidateScorer.LabelProximityBonus(1) + GeneralInvoiceCandidateScorer.DateFormatBonus(nx.Groups(1).Value), "标签下一行")
                End If
            End If
        Next
        Dim g As Match = Regex.Match(compactAll, "开票日期[:：]?" & datePat)
        If g.Success Then
            AddCand(parse, InvoiceFieldNames.InvoiceDate, g.Groups(1).Value, Nothing, -1, "global",
                    2 + GeneralInvoiceCandidateScorer.DateFormatBonus(g.Groups(1).Value), "全局兜底")
        End If
    End Sub

    ''' <summary>购/销名称：定位购买方/销售方锚点，名称候选按 X 就近（有坐标）或行序就近归栏。</summary>
    Private Sub ParseParties(lines As List(Of PdfTextLine), parse As GeneralInvoiceParseResult)
        Dim buyerIdx As Integer = -1, sellerIdx As Integer = -1
        Dim buyerX As Double = Double.NaN, sellerX As Double = Double.NaN

        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If buyerIdx < 0 AndAlso (c.Contains("购买方") OrElse c.Contains("购方")) Then
                buyerIdx = i : buyerX = LineX(lines(i))
            End If
            If sellerIdx < 0 AndAlso (c.Contains("销售方") OrElse c.Contains("销方")) Then
                sellerIdx = i : sellerX = LineX(lines(i))
            End If
        Next

        ' 1) 直接带方位标签的名称（强信号）
        Dim gotSeller As Boolean = False, gotBuyer As Boolean = False
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            Dim ms As Match = Regex.Match(c, "销售方名称[:：]?\s*([^\r\n]{2,40}?)" & NameStop)
            If ms.Success Then
                AddCand(parse, InvoiceFieldNames.SellerName, CleanName(ms.Groups(1).Value), lines(i), i, "labeled-seller",
                        10 + GeneralInvoiceCandidateScorer.OrgNameBonus(ms.Groups(1).Value), "销售方名称标签")
                gotSeller = True
            End If
            Dim mb As Match = Regex.Match(c, "购买方名称[:：]?\s*([^\r\n]{2,40}?)" & NameStop)
            If mb.Success Then
                AddCand(parse, InvoiceFieldNames.BuyerName, CleanName(mb.Groups(1).Value), lines(i), i, "labeled-buyer",
                        10 + GeneralInvoiceCandidateScorer.OrgNameBonus(mb.Groups(1).Value), "购买方名称标签")
                gotBuyer = True
            End If
        Next

        ' 2) 通用「名称:」候选，按锚点归栏（补齐未直接标注的一侧）
        ' 同时处理两栏版式：「销 名称：... 购 名称：...」或「购 名称：... 销 名称：...」
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If c.Contains("销售方名称") OrElse c.Contains("购买方名称") Then Continue For
            ' 排除「项目名称/货物或应税劳务名称」等明细表头，避免把明细列名当作购/销名称
            If c.Contains("项目名称") OrElse c.Contains("货物或应税") OrElse c.Contains("品名") Then Continue For

            ' 先尝试明确带「销 名称」或「购 名称」的两栏版式
            Dim mSeller As Match = Regex.Match(c, "销\s*名称[:：]?\s*([^\r\n]{2,40}?)" & NameStop)
            If mSeller.Success Then
                Dim val As String = CleanName(mSeller.Groups(1).Value)
                If GeneralInvoiceCandidateScorer.OrgNameBonus(val) >= 0 AndAlso Not gotSeller Then
                    AddCand(parse, InvoiceFieldNames.SellerName, val, lines(i), i, "explicit-seller",
                            7 + GeneralInvoiceCandidateScorer.OrgNameBonus(val), "销方名称标签")
                    gotSeller = True
                End If
            End If

            Dim mBuyer As Match = Regex.Match(c, "购\s*名称[:：]?\s*([^\r\n]{2,40}?)" & NameStop)
            If mBuyer.Success Then
                Dim val As String = CleanName(mBuyer.Groups(1).Value)
                If GeneralInvoiceCandidateScorer.OrgNameBonus(val) >= 0 AndAlso Not gotBuyer Then
                    AddCand(parse, InvoiceFieldNames.BuyerName, val, lines(i), i, "explicit-buyer",
                            7 + GeneralInvoiceCandidateScorer.OrgNameBonus(val), "购方名称标签")
                    gotBuyer = True
                End If
            End If

            ' 再用通用「名称:」规则补缺（仅无锚点时）
            If (Not mSeller.Success AndAlso Not mBuyer.Success) Then
                Dim m As Match = Regex.Match(c, "名称[:：]\s*([^\r\n]{2,40}?)" & NameStop)
                If Not m.Success Then Continue For
                Dim val As String = CleanName(m.Groups(1).Value)
                If GeneralInvoiceCandidateScorer.OrgNameBonus(val) < 0 Then Continue For
                Dim side As String = NearestParty(i, LineX(lines(i)), buyerIdx, buyerX, sellerIdx, sellerX)
                If side = "seller" AndAlso Not gotSeller Then
                    AddCand(parse, InvoiceFieldNames.SellerName, val, lines(i), i, "name-anchor-seller",
                            6 + GeneralInvoiceCandidateScorer.OrgNameBonus(val), "名称就近销售方")
                ElseIf side = "buyer" AndAlso Not gotBuyer Then
                    AddCand(parse, InvoiceFieldNames.BuyerName, val, lines(i), i, "name-anchor-buyer",
                            6 + GeneralInvoiceCandidateScorer.OrgNameBonus(val), "名称就近购买方")
                End If
            End If
        Next

        ' 3) 税号（就近归栏；不阻断命名）
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            Dim m As Match = Regex.Match(c, "(?:统一社会信用代码/纳税人识别号|统一社会信用代码|纳税人识别号)[:：]?\s*([0-9A-Z]{15,20})")
            If Not m.Success Then Continue For
            Dim side As String = NearestParty(i, LineX(lines(i)), buyerIdx, buyerX, sellerIdx, sellerX)
            If side = "seller" Then
                AddCand(parse, InvoiceFieldNames.SellerTaxId, m.Groups(1).Value, lines(i), i, "taxid-seller", 5, "税号就近销售方")
            Else
                AddCand(parse, InvoiceFieldNames.BuyerTaxId, m.Groups(1).Value, lines(i), i, "taxid-buyer", 5, "税号就近购买方")
            End If
        Next
    End Sub

    Private Sub ParseTotalWithTax(lines As List(Of PdfTextLine), compactAll As String, parse As GeneralInvoiceParseResult)
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            Dim hasTotalLabel As Boolean = c.Contains("价税合计")
            Dim hasSmall As Boolean = c.Contains("小写")
            If Not hasTotalLabel AndAlso Not hasSmall Then Continue For
            ' 取该行内最后一个两位小数金额（小写通常在大写之后）
            Dim ms As MatchCollection = Regex.Matches(c, "\d+\.\d{2}")
            If ms.Count > 0 Then
                Dim val As String = ms(ms.Count - 1).Value
                AddCand(parse, InvoiceFieldNames.TotalWithTax, val, lines(i), i, "total-label",
                        GeneralInvoiceCandidateScorer.LabelProximityBonus(0) + GeneralInvoiceCandidateScorer.AmountContextBonus(c, hasTotalLabel), "价税合计/小写同行")
            ElseIf i + 1 < lines.Count Then
                Dim nx As MatchCollection = Regex.Matches(Compact(lines(i + 1)), "\d+\.\d{2}")
                If nx.Count > 0 Then
                    AddCand(parse, InvoiceFieldNames.TotalWithTax, nx(nx.Count - 1).Value, lines(i + 1), i + 1, "total-next-line",
                            GeneralInvoiceCandidateScorer.LabelProximityBonus(1) + GeneralInvoiceCandidateScorer.AmountContextBonus(c, hasTotalLabel), "价税合计下一行")
                End If
            End If
        Next
        ' 全局兜底：价税合计 后首个金额
        Dim g As Match = Regex.Match(compactAll, "价税合计[^0-9]{0,24}?[¥￥]?\s*(\d+\.\d{2})")
        If g.Success Then
            AddCand(parse, InvoiceFieldNames.TotalWithTax, g.Groups(1).Value, Nothing, -1, "global", 4, "全局兜底")
        End If
    End Sub

    Private Sub ParseTaxRateAndAmount(lines As List(Of PdfTextLine), compactAll As String, parse As GeneralInvoiceParseResult)
        ' 税额合计：含「税额」且非明细行的金额
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If c.Contains("合计") AndAlso c.Contains("税额") Then
                Dim ms As MatchCollection = Regex.Matches(c, "\d+\.\d{2}")
                If ms.Count > 0 Then
                    AddCand(parse, InvoiceFieldNames.TaxAmount, ms(ms.Count - 1).Value, lines(i), i, "tax-total", 5, "税额合计")
                End If
            End If
        Next
        ' 税率：全局首个（多数常规发票单一税率）
        Dim rate As String = LocalTextNormalizer.NormalizeTaxRate(compactAll)
        If Not String.IsNullOrEmpty(rate) Then
            AddCand(parse, InvoiceFieldNames.TaxRate, rate, Nothing, -1, "global-rate", 3, "全局税率")
        End If
    End Sub

    ' ========== 工具 ==========

    Private Function LineX(ln As PdfTextLine) As Double
        If ln Is Nothing OrElse ln.Words Is Nothing OrElse ln.Words.Count = 0 Then Return Double.NaN
        Return ln.Words(0).X
    End Function

    ''' <summary>把名称候选归到购买方或销售方：有坐标按 X 就近，否则按行序就近。</summary>
    Private Function NearestParty(candIdx As Integer, candX As Double, buyerIdx As Integer, buyerX As Double, sellerIdx As Integer, sellerX As Double) As String
        Dim coords As Boolean = Not Double.IsNaN(candX) AndAlso Not Double.IsNaN(buyerX) AndAlso Not Double.IsNaN(sellerX) AndAlso candX <> 0 AndAlso (buyerX <> 0 OrElse sellerX <> 0) AndAlso buyerX <> sellerX
        If coords Then
            Return If(Math.Abs(candX - sellerX) <= Math.Abs(candX - buyerX), "seller", "buyer")
        End If
        Dim dS As Integer = If(sellerIdx >= 0, Math.Abs(candIdx - sellerIdx), Integer.MaxValue)
        Dim dB As Integer = If(buyerIdx >= 0, Math.Abs(candIdx - buyerIdx), Integer.MaxValue)
        If dS = Integer.MaxValue AndAlso dB = Integer.MaxValue Then Return "seller" ' 无锚点默认销售方（命名更需要）
        Return If(dS <= dB, "seller", "buyer")
    End Function

    Private Function CleanName(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return Nothing
        Dim v As String = s.Trim()
        ' 去掉可能并入的尾随标签片段与分栏符（购/销）
        v = Regex.Replace(v, "[购销]$", "").Trim() ' 清除行尾"购""销"分栏符
        v = Regex.Replace(v, "(地址|开户行|电话|账号|统一社会信用代码|纳税人识别号).*$", "").Trim()
        Return v
    End Function

    Private Sub CheckAmountConsistency(inv As InvoiceInfo, parse As GeneralInvoiceParseResult)
        Try
            If inv.LineItems Is Nothing OrElse inv.LineItems.Count = 0 OrElse String.IsNullOrWhiteSpace(inv.TotalWithTax) Then Return
            Dim sumAmt As Double = 0, sumTax As Double = 0
            ' 必须用 InvariantCulture：发票金额固定是 "1234.56" 形式，在 de-DE 等区域下
            ' 用当前区域解析会把 "138.46" 读成 13846、"1,234.50" 直接解析失败。
            For Each li As InvoiceLineItem In inv.LineItems
                Dim a As Double
                If Double.TryParse(If(li.Amount, ""), NumberStyles.Float, CultureInfo.InvariantCulture, a) Then sumAmt += a
                Dim t As Double
                If Double.TryParse(If(li.TaxAmount, ""), NumberStyles.Float, CultureInfo.InvariantCulture, t) Then sumTax += t
            Next
            Dim total As Double
            If Double.TryParse(inv.TotalWithTax, NumberStyles.Float, CultureInfo.InvariantCulture, total) AndAlso (sumAmt + sumTax) > 0 Then
                If Math.Abs((sumAmt + sumTax) - total) > 0.05 Then
                    parse.Notes.Add(String.Format(CultureInfo.InvariantCulture,
                                                  "明细金额合计({0:F2})+税额({1:F2}) 与价税合计({2:F2}) 不一致（仅告警）。", sumAmt, sumTax, total))
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub AddCand(parse As GeneralInvoiceParseResult, field As String, value As String, ln As PdfTextLine, lineIdx As Integer, rule As String, score As Double, reason As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        Dim c As New GeneralInvoiceFieldCandidate With {
            .FieldName = field, .Value = value.Trim(), .SourceRule = rule, .ConfidenceScore = score, .Reason = reason,
            .LineIndex = lineIdx, .PageNumber = If(ln IsNot Nothing, ln.PageIndex, 0),
            .X = If(ln IsNot Nothing AndAlso ln.Words.Count > 0, ln.Words(0).X, 0),
            .Y = If(ln IsNot Nothing, ln.Y, 0),
            .RawText = If(ln IsNot Nothing, ln.Text, Nothing)}
        parse.AddCandidate(c)
    End Sub

    ''' <summary>把某字段的最优候选写入 result.Fields 并返回其值。</summary>
    Private Function Adopt(result As InvoiceRecognitionResult, parse As GeneralInvoiceParseResult, field As String) As String
        Dim c As GeneralInvoiceFieldCandidate = Nothing
        If parse.Chosen.TryGetValue(field, c) AndAlso c IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(c.Value) Then
            result.Fields.Add(New InvoiceField(field, c.Value, RecognizerLabel))
            Return c.Value
        End If
        Return Nothing
    End Function

End Class
