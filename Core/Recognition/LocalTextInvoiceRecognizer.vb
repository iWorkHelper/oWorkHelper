Imports System.Text.RegularExpressions
Imports System.Collections.Generic

''' <summary>
''' 本地文本识别器：基于 PdfPig 抽取出的文本层做字段解析（正则）。
''' 规则已依据 5 份脱敏滴滴"合并 PDF"样例校准（第 1 页电子发票 + 第 2 页行程单）。
''' 说明：不是 OCR，仅在文本型 PDF 上有效；无有效文本时返回 NeedsOcr。
''' </summary>
Public Class LocalTextInvoiceRecognizer
    Implements IInvoiceRecognizer

    ''' <summary>常规发票专用本地识别器（非滴滴/非行程单发票委派给它）。</summary>
    Private ReadOnly _generalRecognizer As New GeneralInvoiceLocalRecognizer()

    Public ReadOnly Property Name As String Implements IInvoiceRecognizer.Name
        Get
            Return "LocalTextInvoiceRecognizer"
        End Get
    End Property

    Public Function IsAvailable() As Boolean Implements IInvoiceRecognizer.IsAvailable
        Return True
    End Function

    Public Function Recognize(context As RecognitionContext) As InvoiceRecognitionResult Implements IInvoiceRecognizer.Recognize
        Dim result As New InvoiceRecognitionResult()
        result.RecognizerName = Name
        result.Source = RecognitionSource.LocalText

        Try
            Dim rawText As String = If(context Is Nothing, Nothing, context.ExtractedText)
            If String.IsNullOrWhiteSpace(rawText) Then
                Return InvoiceRecognitionResult.NeedsOcr("本地文本层为空，疑似图片型 PDF，需 OCR。", rawText)
            End If

            ' 归一化（全/半角、冒号、空格、金额符号），降低正则失配。
            Dim text As String = LocalTextNormalizer.Normalize(rawText)
            result.RawText = rawText ' 保留原始文本供诊断
            result.DocumentType = DetectDocumentType(text)

            ' 常规发票（非滴滴/非行程单）委派给专用识别器（候选评分 + 分区解析 + 明细）。
            ' 滴滴发票与行程单仍走本类既有逻辑（DocumentType=RideTripStatement）。
            If result.DocumentType = InvoiceDocumentType.VatInvoice Then
                Return _generalRecognizer.Recognize(context)
            End If

            ParseVatInvoiceFields(text, result)
            Dim extract As PdfTextExtractResult = If(context Is Nothing, Nothing, context.TextExtractResult)
            ParseTripFields(text, result, extract)

            Dim invoice As InvoiceInfo = result.Invoice
            result.MissingKeyFields = KeyFieldEvaluator.GetMissingNamingFields(invoice, result.DocumentType)
            If invoice.HasAnyCoreField() Then
                If KeyFieldEvaluator.IsNamingSufficient(invoice, result.DocumentType) Then
                    result.Status = RecognitionStatus.Success
                    result.Message = "本地文本解析成功（命名核心字段齐全）。"
                Else
                    ' 命名核心字段不足（常见：起点/终点未取到）→ 部分成功，供上层回退在线 OCR。
                    result.Status = RecognitionStatus.PartialSuccess
                    result.Message = "本地解析部分成功，缺少命名核心字段：" & String.Join("、", result.MissingKeyFields.ToArray())
                End If
            Else
                Return InvoiceRecognitionResult.NeedsOcr("本地有文本但未解析到关键字段，建议 OCR。", text)
            End If

            AppLogger.Info(String.Format("本地识别: 类型={0}, 状态={1}, 字段数={2}",
                                         result.DocumentType, result.Status, result.Fields.Count))
            Return result

        Catch ex As Exception
            AppLogger.Error("本地文本识别异常。", ex)
            Return InvoiceRecognitionResult.Failure("本地文本识别异常: " & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

    Private Function DetectDocumentType(text As String) As InvoiceDocumentType
        Dim isInvoice As Boolean = text.Contains("发票号码") OrElse text.Contains("电子发票") OrElse text.Contains("增值税")
        ' 收紧滴滴判定：不再“全文出现 行程单/网约车/滴滴 即判定”，避免销售方名称（如“滴滴出行科技有限公司”）
        ' 或备注里偶然提到“行程单”就把常规增值税发票误判为行程单（O-38）。
        Dim isRide As Boolean = HasRideTripEvidence(text)
        ' 合并 PDF 同时含发票与行程单：优先按行程单标记（滴滴场景），发票字段仍会解析。
        If isRide Then
            Return InvoiceDocumentType.RideTripStatement
        End If
        If isInvoice Then
            Return InvoiceDocumentType.VatInvoice
        End If
        Return InvoiceDocumentType.Unknown
    End Function

    ''' <summary>
    ''' 判断是否为网约车/滴滴行程单。要求“行程单专属证据”，而非任意位置出现关键词：
    '''  - 行程明细/表头强信号：上车时间、“共N笔行程”、行程起止、行程明细；
    '''  - 行程单表格表头：同时出现“起点”与“终点”；
    '''  - “行程单/网约车”出现在**表头位置**（独占短行，且不是“备注”字段行）。
    ''' 仅“滴滴”二字（常见于销售方名称或备注）不构成判定依据。
    ''' </summary>
    Private Function HasRideTripEvidence(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False

        ' —— 行程明细/表头强信号 ——
        If text.Contains("上车时间") Then Return True
        If text.Contains("行程起止") Then Return True
        If text.Contains("行程明细") Then Return True
        If Regex.IsMatch(text, "共\s*[0-9]+\s*笔行程") Then Return True
        ' 行程单表格表头：起点 + 终点
        If text.Contains("起点") AndAlso text.Contains("终点") Then Return True

        ' —— “行程单/网约车”需处在表头位置：独占短行且非备注行 ——
        For Each ln As String In Regex.Split(text, "\r?\n")
            Dim t As String = If(ln, "").Trim()
            If t.Length = 0 OrElse t.Length > 24 Then Continue For
            If t.StartsWith("备注", StringComparison.Ordinal) Then Continue For
            If t.Contains("行程单") OrElse t.Contains("网约车") Then Return True
        Next

        Return False
    End Function

    Private Sub ParseVatInvoiceFields(text As String, result As InvoiceRecognitionResult)
        Dim inv As InvoiceInfo = result.Invoice

        inv.InvoiceNumber = SetField(result, InvoiceFieldNames.InvoiceNumber, MatchFirst(text, "发票号码[:：\s]*([0-9]{8,20})"))
        inv.InvoiceCode = SetField(result, InvoiceFieldNames.InvoiceCode, MatchFirst(text, "发票代码[:：\s]*([0-9]{10,12})"))
        inv.InvoiceDate = SetField(result, InvoiceFieldNames.InvoiceDate, MatchFirst(text, "开票日期[:：\s]*([0-9]{4}\s*年\s*[0-9]{1,2}\s*月\s*[0-9]{1,2}\s*日)"))

        ' 价税合计：样例形如"价税合计（大写）（小写）138.46¥"，取 价税合计 后首个金额。
        inv.TotalWithTax = SetField(result, InvoiceFieldNames.TotalWithTax, MatchFirst(text, "价税合计[^0-9]{0,16}?([0-9]+\.[0-9]{2})"))

        ' 税率：样例中金额与税率在文本层相邻（如"134.433%"），取紧邻 % 前的单个数字得到 3%。
        ' 说明：13%/16% 等两位税率在此拼接场景下可能仅取到个位，需 OCR 或后续按样例细化。
        Dim rate As String = MatchFirst(text, "([0-9])\s*%")
        If Not String.IsNullOrEmpty(rate) Then
            inv.TaxRate = SetField(result, InvoiceFieldNames.TaxRate, rate & "%")
        End If

        ' 购/销方名称与纳税人识别号：滴滴发票含两个"名称：XXX统一社会信用代码"。
        AssignParties(text, result)

        inv.ItemName = SetField(result, InvoiceFieldNames.ItemName, MatchFirst(text, "(\*[^*]+\*[^\d\r\n]{2,20})"))
        inv.Drawer = SetField(result, InvoiceFieldNames.Drawer, MatchFirst(text, "开票人[:：\s]*([^\r\n\d]{2,10})"))
    End Sub

    ''' <summary>
    ''' 解析购买方/销售方名称与税号。滴滴场景：名称含"滴滴"者为销售方，另一为购买方。
    ''' 若无"滴滴"标识，按出现顺序回退（第一个名称视为购买方）。
    ''' </summary>
    Private Sub AssignParties(text As String, result As InvoiceRecognitionResult)
        Dim inv As InvoiceInfo = result.Invoice

        Dim names As New List(Of String)()
        For Each m As Match In Regex.Matches(text, "名\s*称[:：]\s*([^：\r\n]+?)(?=统一社会信用代码|纳税人识别号)")
            If m.Groups.Count > 1 Then
                names.Add(m.Groups(1).Value.Trim())
            End If
        Next

        Dim taxIds As New List(Of String)()
        For Each m As Match In Regex.Matches(text, "(?:统一社会信用代码/纳税人识别号|纳税人识别号)[:：]\s*([0-9A-Z]{15,20})")
            If m.Groups.Count > 1 Then
                taxIds.Add(m.Groups(1).Value.Trim())
            End If
        Next

        Dim sellerName As String = Nothing
        Dim buyerName As String = Nothing
        Dim sellerIndex As Integer = -1

        For i As Integer = 0 To names.Count - 1
            If names(i).Contains("滴滴") Then
                sellerName = names(i)
                sellerIndex = i
            End If
        Next

        If sellerName Is Nothing AndAlso names.Count > 0 Then
            ' 回退：末位视为销售方，首位视为购买方（滴滴样例中销售方在后）。
            sellerIndex = names.Count - 1
            sellerName = names(sellerIndex)
        End If

        For i As Integer = 0 To names.Count - 1
            If i <> sellerIndex AndAlso buyerName Is Nothing Then
                buyerName = names(i)
            End If
        Next

        If Not String.IsNullOrEmpty(sellerName) Then inv.SellerName = SetField(result, InvoiceFieldNames.SellerName, sellerName)
        If Not String.IsNullOrEmpty(buyerName) Then inv.BuyerName = SetField(result, InvoiceFieldNames.BuyerName, buyerName)

        ' 税号按名称顺序配对（如数量一致）。
        If taxIds.Count = names.Count AndAlso names.Count > 0 Then
            If sellerIndex >= 0 AndAlso sellerIndex < taxIds.Count Then
                inv.SellerTaxId = SetField(result, InvoiceFieldNames.SellerTaxId, taxIds(sellerIndex))
            End If
            For i As Integer = 0 To taxIds.Count - 1
                If i <> sellerIndex Then
                    inv.BuyerTaxId = SetField(result, InvoiceFieldNames.BuyerTaxId, taxIds(i))
                    Exit For
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 网约车/滴滴行程单解析：表头级字段 + **多条行程明细**。
    ''' 明细行结构（依样例校准）：序号 车型 上车时间 [周X] 城市 起点 终点 里程 金额 [备注]。
    ''' 单条明细解析失败不影响其它明细与整份 PDF。
    ''' </summary>
    Private Sub ParseTripFields(text As String, result As InvoiceRecognitionResult, extract As PdfTextExtractResult)
        If result.DocumentType <> InvoiceDocumentType.RideTripStatement Then
            Return
        End If
        Dim inv As InvoiceInfo = result.Invoice

        ' —— 表头级（归一化文本正则）——
        Dim tripDate As String = MatchFirst(text, "行程起止日期[:]\s*([0-9]{4}-[0-9]{1,2}-[0-9]{1,2})")
        Dim tripTotal As String = MatchFirst(text, "合计\s*([0-9]+\.[0-9]{2})\s*元")
        Dim phone As String = MatchFirst(text, "行程人手机号[:]\s*([0-9]{11})")
        Dim passenger As String = MatchFirst(text, "(?:乘车人|申请人|姓名)[:]\s*([^\r\n:\s]{1,20})")
        Dim statedCount As String = MatchFirst(text, "共\s*([0-9]+)\s*笔行程")

        If Not String.IsNullOrEmpty(tripDate) Then
            inv.SetExtendedField("行程起止日期", tripDate)
            result.Fields.Add(New InvoiceField("行程起止日期", tripDate, Name))
        End If
        If Not String.IsNullOrEmpty(tripTotal) Then
            inv.SetExtendedField(InvoiceFieldNames.TripAmount, tripTotal)
            If String.IsNullOrEmpty(inv.TotalWithTax) Then inv.TotalWithTax = tripTotal
        End If
        If Not String.IsNullOrEmpty(phone) Then inv.SetExtendedField("行程人手机号", phone)
        If Not String.IsNullOrEmpty(statedCount) Then
            Dim n As Integer
            If Integer.TryParse(statedCount, n) Then inv.StatedTripCount = n
        End If

        ' —— 行程明细：优先“坐标行重建按列取单元格”（对地址长短/换行稳），失败回退全文正则 ——
        Dim trip As InvoiceTripInfo = ParseTripTableByLayout(extract)
        If trip Is Nothing Then
            ParseTripRows(text, tripDate, passenger, inv)
        Else
            ' 用表头级信息补全
            If String.IsNullOrEmpty(trip.DepartureTime) Then trip.DepartureTime = tripDate
            If String.IsNullOrEmpty(trip.TripAmount) Then trip.TripAmount = tripTotal
            If String.IsNullOrEmpty(trip.Passenger) Then trip.Passenger = passenger
            inv.Trips.Add(trip)
            result.Fields.Add(New InvoiceField("出发地点", If(trip.StartLocation, ""), Name & ":Layout"))
            result.Fields.Add(New InvoiceField("到达地点", If(trip.EndLocation, ""), Name & ":Layout"))
        End If

        ' 若仍无明细但有表头信息，至少建一条汇总行程，保证命名可用。
        If inv.Trips.Count = 0 AndAlso (Not String.IsNullOrEmpty(tripDate) OrElse Not String.IsNullOrEmpty(tripTotal)) Then
            inv.Trips.Add(New InvoiceTripInfo With {
                .RowIndex = 1, .Passenger = passenger, .DepartureTime = tripDate, .TripAmount = tripTotal})
        End If

        ' 行程金额回填价税合计（供 {金额} 优先取行程金额）
        If inv.Trips.Count > 0 AndAlso Not String.IsNullOrEmpty(inv.Trips(0).TripAmount) AndAlso String.IsNullOrEmpty(inv.TotalWithTax) Then
            inv.TotalWithTax = inv.Trips(0).TripAmount
        End If

        AppLogger.Info("行程单解析：声明笔数=" & inv.StatedTripCount & "，实际明细=" & inv.Trips.Count &
                       "，起点=" & TripFieldLog(inv, Function(t) t.StartLocation) & "，终点=" & TripFieldLog(inv, Function(t) t.EndLocation))

        ' O-39：多行程单当前只把“首条行程”用于命名（起点/终点/金额），表头级合计金额仍完整。
        ' 声明笔数与实际解析条数不一致时明确写入结果消息并告警，避免误以为已覆盖全部行程。
        If inv.StatedTripCount > inv.Trips.Count Then
            Dim mismatch As String = "行程单声明 " & inv.StatedTripCount & " 笔行程，实际仅解析出 " & inv.Trips.Count &
                                     " 笔；命名使用首条行程的起点/终点/金额（多行程限制见 docs\TROUBLESHOOTING.md）。"
            result.Messages.Add(mismatch)
            AppLogger.Warn(mismatch)
        End If
    End Sub

    Private Function TripFieldLog(inv As InvoiceInfo, getter As Func(Of InvoiceTripInfo, String)) As String
        If inv.Trips Is Nothing OrElse inv.Trips.Count = 0 Then Return "(无)"
        Return If(getter(inv.Trips(0)), "(空)")
    End Function

    ''' <summary>
    ''' 坐标行重建按列解析行程单表格。返回首条行程；无表格/失败返回 Nothing。
    ''' 稳健点：读取表头列 X（每份 PDF 列宽不同），数据词按最近列中心归列，支持地址换行（同列多行拼接）。
    ''' </summary>
    Private Function ParseTripTableByLayout(extract As PdfTextExtractResult) As InvoiceTripInfo
        Try
            If extract Is Nothing OrElse extract.Lines Is Nothing OrElse extract.Lines.Count = 0 Then
                Return Nothing
            End If

            ' 定位表头行：同一行含“起点”和“终点”。
            Dim header As PdfTextLine = Nothing
            For Each ln As PdfTextLine In extract.Lines
                If ln.FindWord("起点") IsNot Nothing AndAlso ln.FindWord("终点") IsNot Nothing Then
                    header = ln
                    Exit For
                End If
            Next
            If header Is Nothing Then Return Nothing

            Dim colStart As PdfTextWord = header.FindWord("起点")
            Dim colEnd As PdfTextWord = header.FindWord("终点")
            Dim colCity As PdfTextWord = header.FindWord("城市")
            Dim colAmount As PdfTextWord = FindWordStartsWith(header, "金额")
            Dim colMileage As PdfTextWord = FindWordStartsWith(header, "里程")

            Dim cxStart As Double = colStart.CenterX
            Dim cxEnd As Double = colEnd.CenterX
            ' 起点/终点 列区间：左界=城市与起点中点，右界=里程列左缘。区间外的词（车型/城市/时间/里程/金额/备注）不归起终点。
            Dim leftBound As Double = If(colCity IsNot Nothing, (colCity.CenterX + cxStart) / 2.0, cxStart - 18)
            Dim rightBound As Double = If(colMileage IsNot Nothing, colMileage.X - 2, cxEnd + 80)

            ' 数据带：表头下方约 45 单位内（覆盖主行 + 一条换行）。
            Dim bandTop As Double = header.Y - 5
            Dim bandBottom As Double = header.Y - 48
            Dim dataWords As New List(Of PdfTextWord)()
            For Each ln As PdfTextLine In extract.Lines
                If ln.PageIndex <> header.PageIndex Then Continue For
                If ln.Y < bandTop AndAlso ln.Y > bandBottom Then
                    dataWords.AddRange(ln.Words)
                End If
            Next
            If dataWords.Count = 0 Then Return Nothing

            Dim trip As New InvoiceTripInfo With {.RowIndex = 1}
            Dim startWords As New List(Of PdfTextWord)()
            Dim endWords As New List(Of PdfTextWord)()

            For Each w As PdfTextWord In dataWords
                Dim t As String = If(w.Text, "").Trim()
                If t.Length = 0 Then Continue For

                ' 备注/序号/车型/时间/城市/里程/金额 用模式识别，避免混入起终点。
                If Regex.IsMatch(t, "^\d+\.\d{2}$") Then
                    ' 金额（两位小数）——取最接近金额列的
                    If String.IsNullOrEmpty(trip.TripAmount) OrElse NearHeader(w, colAmount) Then trip.TripAmount = t
                    Continue For
                End If
                If Regex.IsMatch(t, "^\d+\.\d$") Then
                    trip.Mileage = t : Continue For
                End If
                If Regex.IsMatch(t, "^\d+$") Then Continue For ' 序号
                If Regex.IsMatch(t, "^(快车|专车|优享|拼车|出租车|豪华车|礼橙专车|滴滴快车)$") Then
                    trip.ServiceType = t : Continue For
                End If
                If Regex.IsMatch(t, "^\d{2}-\d{2}$") OrElse Regex.IsMatch(t, "^\d{2}:\d{2}$") OrElse Regex.IsMatch(t, "^周[一二三四五六日]$") Then
                    If String.IsNullOrEmpty(trip.DepartureTime) Then trip.DepartureTime = t Else trip.DepartureTime &= " " & t
                    Continue For
                End If
                If t.EndsWith("市", StringComparison.Ordinal) AndAlso t.Length <= 6 Then
                    trip.City = t : Continue For
                End If
                If t.Contains("企业打车") OrElse t.StartsWith("个人", StringComparison.Ordinal) Then Continue For
                If t.EndsWith("市", StringComparison.Ordinal) AndAlso t.Length <= 3 Then Continue For ' 城市残词（如“市”）

                ' 只保留落在 起点/终点 列区间内的词（排除车型/城市/时间/里程/金额/备注列）。
                If w.CenterX < leftBound OrElse w.CenterX >= rightBound Then Continue For

                ' 归列：起点/终点 按最近列中心。
                Dim dStart As Double = Math.Abs(w.CenterX - cxStart)
                Dim dEnd As Double = Math.Abs(w.CenterX - cxEnd)
                If dStart <= dEnd Then startWords.Add(w) Else endWords.Add(w)
            Next

            trip.StartLocation = JoinCell(startWords)
            trip.EndLocation = JoinCell(endWords)

            ' 无有效起终点与金额 → 视为解析失败
            If String.IsNullOrEmpty(trip.StartLocation) AndAlso String.IsNullOrEmpty(trip.EndLocation) AndAlso String.IsNullOrEmpty(trip.TripAmount) Then
                Return Nothing
            End If
            Return trip
        Catch ex As Exception
            AppLogger.Warn("坐标表格解析异常（回退正则）：" & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function NearHeader(w As PdfTextWord, header As PdfTextWord) As Boolean
        If header Is Nothing Then Return False
        Return Math.Abs(w.CenterX - header.CenterX) < 30
    End Function

    Private Function FindWordStartsWith(line As PdfTextLine, prefix As String) As PdfTextWord
        For Each w As PdfTextWord In line.Words
            If w.Text IsNot Nothing AndAlso w.Text.StartsWith(prefix, StringComparison.Ordinal) Then Return w
        Next
        Return Nothing
    End Function

    ''' <summary>把同列多个词（可能换行）按 Y 从上到下、X 从左到右拼接为单元格文本。</summary>
    Private Function JoinCell(words As List(Of PdfTextWord)) As String
        If words Is Nothing OrElse words.Count = 0 Then Return Nothing
        Dim sorted As New List(Of PdfTextWord)(words)
        sorted.Sort(Function(a, b)
                        Dim c As Integer = b.Y.CompareTo(a.Y) ' Y 大在前（靠上）
                        If c <> 0 Then Return c
                        Return a.X.CompareTo(b.X)
                    End Function)
        Dim sb As New System.Text.StringBuilder()
        For Each w As PdfTextWord In sorted
            sb.Append(w.Text)
        Next
        Return sb.ToString().Trim()
    End Function

    ''' <summary>解析明细行（可多条）。以"上车时间"为锚点分段。</summary>
    Private Sub ParseTripRows(text As String, tripDate As String, passenger As String, inv As InvoiceInfo)
        Try
            ' 明细区通常在"备注"表头之后；找不到则对全文尝试。
            Dim area As String = text
            Dim headerIdx As Integer = text.IndexOf("金额")
            If headerIdx >= 0 Then
                area = text.Substring(headerIdx)
            End If

            ' 行模式：序号 车型 上车时间(MM-DD HH:MM) [周X] 城市(..市) 起终点 里程(x.x) 金额(x.xx) [备注]
            Dim pattern As String =
                "([0-9]{1,3})(快车|专车|优享|拼车|出租车|豪华车|礼橙专车|滴滴快车)" &
                "([0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2})\s*(?:周[一二三四五六日])?" &
                "([一-龥]{2,8}?市)?(.+?)([0-9]+\.[0-9])([0-9]+\.[0-9]{2})(企业打车|个人[^0-9]{0,6})?"

            Dim matches As MatchCollection = Regex.Matches(area, pattern)
            Dim seq As Integer = 0
            For Each m As Match In matches
                seq += 1
                Try
                    Dim trip As New InvoiceTripInfo()
                    trip.RowIndex = ToInt(m.Groups(1).Value, seq)
                    trip.ServiceType = NullIfEmpty(m.Groups(2).Value)
                    trip.DepartureTime = NullIfEmpty(m.Groups(3).Value.Trim())
                    trip.City = NullIfEmpty(m.Groups(4).Value)
                    trip.Passenger = passenger
                    trip.Mileage = NullIfEmpty(m.Groups(6).Value)
                    trip.TripAmount = NullIfEmpty(m.Groups(7).Value)
                    SplitStartEnd(m.Groups(5).Value, trip)
                    inv.Trips.Add(trip)
                Catch
                    ' 单行失败跳过。
                End Try
            Next
        Catch ex As Exception
            AppLogger.Warn("行程明细解析异常（忽略，不影响整体）：" & ex.Message)
        End Try
    End Sub

    ''' <summary>尽力把"起点终点"合并串拆分。样例形如"大道|飞行者大会(观众入口)对面东西湖区|洺悦御府-北门"。</summary>
    Private Sub SplitStartEnd(mid As String, trip As InvoiceTripInfo)
        If String.IsNullOrWhiteSpace(mid) Then Return
        mid = mid.Trim()
        ' 起点/终点常各含一个"区|地点"结构，尝试按第二个"区名|"边界切分。
        Dim m As Match = Regex.Match(mid, "^(.*?[\|｜].*?)([一-龥]{2,6}区[\|｜].*)$")
        If m.Success Then
            trip.StartLocation = m.Groups(1).Value.Trim()
            trip.EndLocation = m.Groups(2).Value.Trim()
        Else
            ' 无法可靠拆分：整体作为起点保留原文，避免信息丢失。
            trip.StartLocation = mid
        End If
    End Sub

    Private Function ToInt(s As String, fallback As Integer) As Integer
        Dim n As Integer
        If Integer.TryParse(s, n) Then Return n
        Return fallback
    End Function

    Private Function NullIfEmpty(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return Nothing
        Return s.Trim()
    End Function

    Private Function SetField(result As InvoiceRecognitionResult, fieldName As String, value As String) As String
        If Not String.IsNullOrWhiteSpace(value) Then
            Dim trimmed As String = value.Trim()
            result.Fields.Add(New InvoiceField(fieldName, trimmed, Name))
            Return trimmed
        End If
        Return Nothing
    End Function

    Private Function MatchFirst(text As String, pattern As String) As String
        Try
            Dim m As Match = Regex.Match(text, pattern)
            If m.Success AndAlso m.Groups.Count > 1 Then
                Return m.Groups(1).Value.Trim()
            End If
        Catch ex As Exception
            ' 正则异常（如模式错误）不应静默吞掉：记日志以便诊断识别率下降的原因。
            AppLogger.Warn("本地识别正则匹配异常（已跳过该模式）：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
        Return Nothing
    End Function

End Class
