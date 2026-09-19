Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' 离线测试入口（不依赖 Outlook / 邮件）。
''' 覆盖：读取 PDF → PdfPig 抽取 → 本地字段解析 →（可选）百度 OCR → 字段映射 → 命名预览 → 摘要。
''' AK/SK 通过环境变量传入（不写死、不打印完整值）：
'''   set BAIDU_OCR_AK=你的AK
'''   set BAIDU_OCR_SK=你的SK
''' </summary>
Module OfflineTesterProgram

    Sub Main(rawArgs As String())
        Try
            AppLogger.Initialize()

            ' —— 资源管理器目录打开/激活诊断（真实环境验证匹配逻辑）——
            If rawArgs IsNot Nothing AndAlso rawArgs.Length >= 1 Then
                Select Case rawArgs(0).ToLowerInvariant()
                    Case "--explorer-scan"
                        Dim target As String = If(rawArgs.Length >= 2, rawArgs(1), Environment.CurrentDirectory)
                        Console.WriteLine(ExplorerFolderService.BuildScanReport(target))
                        Return
                    Case "--explorer-normalize"
                        If rawArgs.Length >= 2 Then
                            For i As Integer = 1 To rawArgs.Length - 1
                                Console.WriteLine(rawArgs(i) & "  =>  " & ExplorerFolderService.NormalizeForDiagnostics(rawArgs(i)))
                            Next
                        End If
                        Return
                    Case "--explorer-open"
                        Dim target As String = If(rawArgs.Length >= 2, rawArgs(1), Environment.CurrentDirectory)
                        ExplorerFolderService.OpenOrActivateFolder(target)
                        Console.WriteLine("已调用 OpenOrActivateFolder（详见日志）：" & target)
                        Return
                End Select
            End If

            Dim opts As CliOptions = CliOptions.Parse(rawArgs)

            If Not opts.ErrorMessage Is Nothing Then
                Console.WriteLine("参数错误: " & opts.ErrorMessage)
                Console.WriteLine("")
                PrintHelp()
                Environment.ExitCode = 2
                Return
            End If

            If opts.SelfTest Then
                SelfTestRunner.Run()
                Return
            End If

            If opts.SaveBaiduConfig Then
                SaveBaiduConfigFromEnv()
                Return
            End If

            If opts.MergeMode Then
                RunMerge(opts.Positionals)
                Return
            End If

            If Not String.IsNullOrEmpty(opts.SimulateError) Then
                RunSimulateError(opts.SimulateError)
                Return
            End If

            If opts.ValidateConfig OrElse opts.Preflight Then
                RunPreflight(opts)
                Return
            End If

            If opts.GeneralInvoice OrElse Not String.IsNullOrEmpty(opts.CompareExpectedPath) Then
                RunGeneralInvoiceDiagnostics(opts)
                Return
            End If

            If Not String.IsNullOrEmpty(opts.ParseOcrJsonPath) Then
                ParseOcrJsonFile(opts.ParseOcrJsonPath)
                Return
            End If

            If opts.ShowHelp OrElse String.IsNullOrEmpty(opts.TargetPath) Then
                PrintHelp()
                Return
            End If

            Dim options As BaiduOcrOptions = BuildOptions(opts)
            Console.WriteLine("模式: " & DescribeMode(opts, options))

            If Not String.IsNullOrEmpty(opts.SaveResponseDir) Then
                Directory.CreateDirectory(opts.SaveResponseDir)
                Console.WriteLine("脱敏 OCR 响应保存目录: " & opts.SaveResponseDir)
            End If

            Dim files As List(Of String) = ResolveFiles(opts.TargetPath)
            If files.Count = 0 Then
                Console.WriteLine("未找到 PDF: " & opts.TargetPath)
                Return
            End If

            For Each pdf As String In files
                ProcessOne(pdf, options, opts)
            Next

        Catch ex As Exception
            Console.WriteLine("运行异常: " & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Sub

    Private Sub ProcessOne(pdf As String, options As BaiduOcrOptions, opts As CliOptions)
        Console.WriteLine("")
        Console.WriteLine("==================================================")
        Console.WriteLine("文件: " & Path.GetFileName(pdf))

        Dim extractor As New PdfTextExtractor()
        Dim extract As PdfTextExtractResult = extractor.Extract(pdf)
        Console.WriteLine(String.Format("  抽取: 成功={0}, 页数={1}, 有效字符={2}, 疑似图片型={3}",
                                        extract.Success, extract.PageCount, extract.TextLength, extract.IsLikelyImageBased))

        Dim pipeline As New RecognitionPipeline(options)
        Dim recog As InvoiceRecognitionResult = pipeline.Recognize(pdf, Path.GetFileName(pdf))

        Console.WriteLine(String.Format("  识别: 状态={0}, 来源={1}, 类型={2}, 字段数={3}",
                                        recog.Status, recog.Source, recog.DocumentType, recog.Fields.Count))
        If Not String.IsNullOrEmpty(recog.Message) Then
            Console.WriteLine("  说明: " & recog.Message)
        End If
        If recog.Status = RecognitionStatus.Failure Then
            Console.WriteLine("  失败原因: " & If(recog.Message, "(未知)"))
        End If

        Dim inv As InvoiceInfo = recog.Invoice
        PrintField("发票号码", inv.InvoiceNumber)
        PrintField("开票日期", inv.InvoiceDate)
        PrintField("销售方名称", inv.SellerName)
        PrintField("购买方名称", inv.BuyerName)
        PrintField("价税合计", inv.TotalWithTax)
        PrintField("税率", inv.TaxRate)

        Dim tripCount As Integer = If(inv.Trips IsNot Nothing, inv.Trips.Count, 0)
        Console.WriteLine("  行程明细数量: " & tripCount & If(inv.StatedTripCount > 0, "（声明 " & inv.StatedTripCount & " 笔）", ""))
        If inv.Trips IsNot Nothing Then
            Dim idx As Integer = 0
            For Each t As InvoiceTripInfo In inv.Trips
                idx += 1
                Console.WriteLine(String.Format("    行程#{0}: 车型={1}, 出发={2}, 城市={3}, 起点={4}, 终点={5}, 里程={6}, 金额={7}",
                                                idx, Dash(t.ServiceType), Dash(t.DepartureTime), Dash(t.City),
                                                Dash(t.StartLocation), Dash(t.EndLocation), Dash(t.Mileage), Dash(t.TripAmount)))
            Next
        End If

        If recog.MissingKeyFields IsNot Nothing AndAlso recog.MissingKeyFields.Count > 0 Then
            Console.WriteLine("  缺失关键字段: " & String.Join("、", recog.MissingKeyFields.ToArray()))
        End If

        ' PDF 分类（滴滴发票/滴滴行程单/常规发票/未识别）
        Dim classifier As New MailProcessingClassifier()
        Dim pdfClass As PdfAttachmentClassification = classifier.ClassifyPdf(recog, recog.RawText, Path.GetFileName(pdf), Nothing)
        Console.WriteLine("  PDF 分类: " & pdfClass.ToString())
        If pdfClass = PdfAttachmentClassification.UnknownPdf Then
            Console.WriteLine("  未识别命名: " & UnknownPdfNamingRule.BuildFileName(Path.GetFileName(pdf)))
        ElseIf pdfClass = PdfAttachmentClassification.GeneralInvoicePdf Then
            Dim gp As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv, recog.DocumentType, Path.GetFileName(pdf), "20260101000000", Nothing, 1, recog.Source.ToString(), NamingTemplates.DefaultGeneralInvoiceTemplate)
            Console.WriteLine("  常规发票命名: " & gp.FileName)
        End If

        ' 命名预览（可用 --template 覆盖统一命名模板做验收）
        Dim rule As ArchiveNamingRule
        If Not String.IsNullOrEmpty(opts.InvoiceTemplateOverride) Then
            Dim tpls As New NamingTemplates()
            tpls.UnifiedTemplate = opts.InvoiceTemplateOverride
            rule = New ArchiveNamingRule(tpls)
        Else
            rule = New ArchiveNamingRule()
        End If
        Dim plan As ArchiveNamePlan = rule.BuildPlan(inv, recog.DocumentType, Path.GetFileName(pdf), "20260101000000", Nothing, 0, recog.Source.ToString())
        Console.WriteLine("  命名规则: " & plan.RuleName & If(plan.FallbackTriggered, "(fallback)", ""))
        Console.WriteLine("  命名预览: " & plan.FileName)
        If plan.UnknownPlaceholders IsNot Nothing AndAlso plan.UnknownPlaceholders.Count > 0 Then
            Console.WriteLine("  未知占位符: " & String.Join("、", plan.UnknownPlaceholders.ToArray()))
        End If

        ' 本地诊断转储（可能含票据信息 → 仅输出到忽略目录 local-debug，勿提交）
        If opts.DumpLocalDebug Then
            DumpLocalDebugFile(opts.LocalDebugDir, pdf, extract, recog)
        End If

        ' 保存脱敏 OCR 响应（仅在线来源且有 JSON 时）
        If Not String.IsNullOrEmpty(opts.SaveResponseDir) _
           AndAlso (recog.Source = RecognitionSource.BaiduOcr OrElse recog.Source = RecognitionSource.Mixed) _
           AndAlso Not String.IsNullOrEmpty(recog.RawText) Then
            SaveDesensitizedResponse(opts.SaveResponseDir, Path.GetFileName(pdf), recog.RawText)
        End If
    End Sub

    ''' <summary>本地识别诊断转储：原始/归一化文本、坐标行、字段结果、命名、缺失、是否需 OCR。</summary>
    Private Sub DumpLocalDebugFile(dir As String, pdf As String, extract As PdfTextExtractResult, recog As InvoiceRecognitionResult)
        Try
            Directory.CreateDirectory(dir)
            Dim sb As New StringBuilder()
            sb.AppendLine("文件: " & Path.GetFileName(pdf))
            sb.AppendLine("页数: " & extract.PageCount & "  有效字符: " & extract.TextLength & "  疑似图片型: " & extract.IsLikelyImageBased)
            sb.AppendLine("")
            sb.AppendLine("===== 原始抽取文本 (page.Text) =====")
            sb.AppendLine(If(extract.Text, ""))
            sb.AppendLine("")
            sb.AppendLine("===== 归一化文本 =====")
            sb.AppendLine(LocalTextNormalizer.Normalize(If(extract.Text, "")))
            sb.AppendLine("")
            sb.AppendLine("===== 坐标行重建（页/Y/词） =====")
            If extract.Lines IsNot Nothing Then
                For Each ln As PdfTextLine In extract.Lines
                    sb.AppendLine(String.Format("P{0} Y={1,7:N1}  {2}", ln.PageIndex, ln.Y, ln.Text))
                Next
            End If
            sb.AppendLine("")
            sb.AppendLine("===== 识别结果 =====")
            sb.AppendLine("状态=" & recog.Status & " 来源=" & recog.Source & " 类型=" & recog.DocumentType)
            Dim inv As InvoiceInfo = recog.Invoice
            sb.AppendLine("发票号码=" & Dash(inv.InvoiceNumber) & " 开票日期=" & Dash(inv.InvoiceDate) & " 价税合计=" & Dash(inv.TotalWithTax))
            If inv.Trips IsNot Nothing AndAlso inv.Trips.Count > 0 Then
                Dim t = inv.Trips(0)
                sb.AppendLine("行程: 出发=" & Dash(t.DepartureTime) & " 金额=" & Dash(t.TripAmount) & " 起点=" & Dash(t.StartLocation) & " 终点=" & Dash(t.EndLocation))
            End If
            sb.AppendLine("命名核心缺失: " & If(recog.MissingKeyFields IsNot Nothing, String.Join("、", recog.MissingKeyFields.ToArray()), ""))
            Dim plan As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv, recog.DocumentType, Path.GetFileName(pdf), "20260101000000", Nothing, 0, recog.Source.ToString())
            sb.AppendLine("命名结果: " & plan.FileName & "  规则=" & plan.RuleName & "  fallback=" & plan.FallbackTriggered)
            sb.AppendLine("是否应回退在线OCR: " & (recog.Status <> RecognitionStatus.Success))

            Dim outPath As String = Path.Combine(dir, Path.GetFileNameWithoutExtension(pdf) & ".localdebug.txt")
            File.WriteAllText(outPath, sb.ToString(), New UTF8Encoding(False))
            Console.WriteLine("  本地诊断已写入: " & outPath & "（含票据信息，勿提交）")
        Catch ex As Exception
            Console.WriteLine("  本地诊断转储失败: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveDesensitizedResponse(dir As String, pdfName As String, rawJson As String)
        Try
            Dim safe As String = Desensitize(rawJson)
            Dim outPath As String = Path.Combine(dir, Path.GetFileNameWithoutExtension(pdfName) & ".ocr.desensitized.json")
            File.WriteAllText(outPath, safe, New UTF8Encoding(False))
            Console.WriteLine("  已保存脱敏 OCR 响应: " & outPath)
        Catch ex As Exception
            Console.WriteLine("  保存脱敏响应失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>脱敏：掩去长数字串（税号/手机号/发票号等）中间部分。</summary>
    Private Function Desensitize(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        ' 11 位及以上数字串 → 保留前2后2
        Dim r1 As String = Regex.Replace(text, "\d{11,}", Function(m) MaskDigits(m.Value))
        ' 15-20 位纳税人识别号（数字+大写字母）
        Dim r2 As String = Regex.Replace(r1, "[0-9A-Z]{15,20}", Function(m) MaskDigits(m.Value))
        Return r2
    End Function

    Private Function MaskDigits(s As String) As String
        If s.Length <= 4 Then Return New String("*"c, s.Length)
        Return s.Substring(0, 2) & New String("*"c, s.Length - 4) & s.Substring(s.Length - 2)
    End Function

    Private Function Dash(s As String) As String
        Return If(String.IsNullOrWhiteSpace(s), "-", s)
    End Function

    Private Sub PrintField(name As String, value As String)
        If Not String.IsNullOrWhiteSpace(value) Then
            Console.WriteLine("    " & name & ": " & value)
        End If
    End Sub

    Private Function BuildOptions(opts As CliOptions) As BaiduOcrOptions
        Dim o As New BaiduOcrOptions()
        If opts.LocalOnly Then
            o.Enabled = False
            o.PreferLocalParse = True
            Return o
        End If

        ' 从本机加密配置文件读取（DPAPI 解密 SK），验证“配置落地→读取→真实 OCR”闭环。
        If opts.UseConfig Then
            Dim fromCfg As BaiduOcrOptions = BaiduXmlConfigStore.Load()
            If fromCfg IsNot Nothing Then
                If opts.ForceOcr Then fromCfg.PreferLocalParse = False
                Return fromCfg
            End If
            Console.WriteLine("未找到本机加密配置文件：" & BaiduXmlConfigStore.GetPath())
        End If

        If opts.UseOcr OrElse opts.ForceOcr Then
            o.Enabled = True
            o.ApiKey = If(Environment.GetEnvironmentVariable("BAIDU_OCR_AK"), "")
            o.SecretKey = If(Environment.GetEnvironmentVariable("BAIDU_OCR_SK"), "")
            o.AutoFallbackToOcr = True
            ' 强制 OCR：以在线为主（本地充分也调用 OCR）。
            o.PreferLocalParse = Not opts.ForceOcr
        Else
            o.Enabled = False
            o.PreferLocalParse = True
        End If
        Return o
    End Function

    ''' <summary>
    ''' 从环境变量 BAIDU_OCR_AK/SK 读取密钥，DPAPI 加密 SK 后写入本机配置文件（Enabled=true）。
    ''' 用于把开发期 OCR 配置安全落地到本机；不打印完整 AK/SK。
    ''' </summary>
    Private Sub SaveBaiduConfigFromEnv()
        Dim ak As String = If(Environment.GetEnvironmentVariable("BAIDU_OCR_AK"), "").Trim()
        Dim sk As String = If(Environment.GetEnvironmentVariable("BAIDU_OCR_SK"), "").Trim()
        If String.IsNullOrEmpty(ak) OrElse String.IsNullOrEmpty(sk) Then
            Console.WriteLine("请先设置环境变量 BAIDU_OCR_AK / BAIDU_OCR_SK 再执行 --save-baidu-config")
            Environment.ExitCode = 2
            Return
        End If
        Dim o As New BaiduOcrOptions()
        o.Enabled = True
        o.ApiKey = ak
        o.PreferLocalParse = True
        o.AutoFallbackToOcr = True
        If Not BaiduXmlConfigStore.Save(o, sk) Then
            Console.WriteLine("保存失败：Secret Key 无法 DPAPI 加密，已拒绝明文落盘。")
            Environment.ExitCode = 1
            Return
        End If
        Console.WriteLine("已将百度 OCR 配置安全写入本机：" & BaiduXmlConfigStore.GetPath())
        Console.WriteLine("  Enabled=true, ApiKey=" & BaiduOcrOptions.MaskSecret(ak) & ", SecretKey=已 DPAPI 加密（DPAPI: 前缀）")
        Console.WriteLine("  提示：该文件已被 .gitignore 忽略，切勿提交。")
    End Sub

    Private Function DescribeMode(opts As CliOptions, o As BaiduOcrOptions) As String
        If opts.LocalOnly Then Return "仅本地解析"
        Dim src As String = If(opts.UseConfig, "（来自本机加密配置）", "")
        If opts.ForceOcr Then Return "强制在线 OCR" & src & OcrCfgNote(o)
        If opts.UseOcr Then Return "本地优先 + 在线兜底" & src & OcrCfgNote(o)
        If opts.UseConfig AndAlso o.IsConfigured() Then Return "使用本机加密配置" & OcrCfgNote(o)
        Return "仅本地解析（默认）"
    End Function

    Private Function OcrCfgNote(o As BaiduOcrOptions) As String
        If o.IsConfigured() Then
            Return "（AK=" & BaiduOcrOptions.MaskSecret(o.ApiKey) & "）"
        End If
        Return "（未配置 AK/SK，将退回本地）"
    End Function

    ''' <summary>展示某错误码（或全部）的用户友好文案（--simulate-error &lt;码名或 list&gt;）。</summary>
    Private Sub RunSimulateError(name As String)
        If String.Equals(name, "list", StringComparison.OrdinalIgnoreCase) Then
            Console.WriteLine("可用错误码（--simulate-error <码名>）：")
            For Each code As AppErrorCode In [Enum].GetValues(GetType(AppErrorCode))
                Console.WriteLine("  " & code.ToString())
            Next
            Return
        End If
        Dim parsed As AppErrorCode
        If Not [Enum].TryParse(name, True, parsed) Then
            Console.WriteLine("未知错误码: " & name & "（用 --simulate-error list 查看全部）")
            Environment.ExitCode = 2
            Return
        End If
        Dim e As AppError = UserFriendlyMessageProvider.Describe(parsed)
        Console.WriteLine("错误码: " & parsed.ToString() & "  级别: " & e.Severity.ToString())
        Console.WriteLine("用户提示: " & e.UserMessage)
        Console.WriteLine("建议操作: " & If(String.IsNullOrEmpty(e.Suggestion), "(无)", e.Suggestion))
    End Sub

    ''' <summary>
    ''' 归档前预检查/配置校验（--preflight / --validate-config）。
    ''' 归档目录取第一个位置参数（或不传→用未配置演示）；OCR 从环境变量；模板可 --template 覆盖。
    ''' </summary>
    Private Sub RunPreflight(opts As CliOptions)
        AppLogger.Initialize()
        Dim archiveFolder As String = If(opts.Positionals.Count > 0, opts.Positionals(0), Nothing)
        Dim o As New BaiduOcrOptions()
        o.ApiKey = If(Environment.GetEnvironmentVariable("BAIDU_OCR_AK"), "")
        o.SecretKey = If(Environment.GetEnvironmentVariable("BAIDU_OCR_SK"), "")
        o.Enabled = Not String.IsNullOrEmpty(o.ApiKey)
        Dim template As String = If(opts.InvoiceTemplateOverride, "{乘车日期}_{金额}_{出发地点}_{到达地点}")

        Dim checker As New ArchivePreflightChecker()
        ' selectedCount 用 1 模拟“已选中邮件”（离线无 Outlook）
        Dim r As ArchivePreflightResult = checker.Check(1, archiveFolder, o, template)
        Console.WriteLine("预检查：问题数=" & r.Issues.Count & "，是否阻断=" & r.HasBlocking)
        Console.WriteLine(r.BuildUserText())
        If r.HasBlocking Then Environment.ExitCode = 1
    End Sub

    ''' <summary>
    ''' 常规发票本地识别诊断（--general-invoice）：分类 / 发票类型 / 逐页行 / 字段候选 / 择优 / 明细 / 缺失 / 是否需 OCR / 命名预览。
    ''' 可选 --compare-expected 读取期望文件做回归对比（不提交真实样例与期望）。
    ''' </summary>
    Private Sub RunGeneralInvoiceDiagnostics(opts As CliOptions)
        AppLogger.Initialize()
        Dim expected As Dictionary(Of String, Dictionary(Of String, String)) = Nothing
        If Not String.IsNullOrEmpty(opts.CompareExpectedPath) Then
            expected = LoadExpected(opts.CompareExpectedPath)
            If expected Is Nothing Then Environment.ExitCode = 2 : Return
            Console.WriteLine("已加载期望文件：" & opts.CompareExpectedPath & "（条目 " & expected.Count & "）")
        End If

        Dim files As List(Of String) = If(String.IsNullOrEmpty(opts.TargetPath), New List(Of String)(), ResolveFiles(opts.TargetPath))
        If files.Count = 0 Then
            Console.WriteLine("未找到 PDF：" & If(opts.TargetPath, "(未指定 pdf 或目录)"))
            Console.WriteLine("用法：OfflineTester.exe <pdf或目录> --general-invoice [--dump-layout --dump-candidates --dump-line-items] [--compare-expected 期望.csv|json]")
            Environment.ExitCode = 2
            Return
        End If

        Dim pass As Integer = 0, fail As Integer = 0
        For Each pdf As String In files
            Dim r As Integer() = DiagnoseGeneralInvoice(pdf, opts, expected)
            pass += r(0) : fail += r(1)
        Next
        If expected IsNot Nothing Then
            Console.WriteLine("")
            Console.WriteLine(String.Format("回归对比合计：字段通过 {0}，失败 {1}", pass, fail))
            If fail > 0 Then Environment.ExitCode = 1
        End If
    End Sub

    ''' <summary>诊断单个 PDF；返回 {通过字段数, 失败字段数}（无期望时为 0,0）。</summary>
    Private Function DiagnoseGeneralInvoice(pdf As String, opts As CliOptions, expected As Dictionary(Of String, Dictionary(Of String, String))) As Integer()
        Console.WriteLine("")
        Console.WriteLine("==================================================")
        Console.WriteLine("文件: " & Path.GetFileName(pdf))

        Dim extractor As New PdfTextExtractor()
        Dim extract As PdfTextExtractResult = extractor.Extract(pdf)
        Dim context As New RecognitionContext With {
            .PdfPath = pdf, .OriginalFileName = Path.GetFileName(pdf),
            .ExtractedText = extract.Text, .TextExtractResult = extract}

        Dim gi As New GeneralInvoiceLocalRecognizer()
        Dim parse As GeneralInvoiceParseResult = gi.Diagnose(context)
        Dim res As InvoiceRecognitionResult = gi.Recognize(context)

        ' 分类
        Dim cls As PdfAttachmentClassification = New MailProcessingClassifier().ClassifyPdf(res, res.RawText, Path.GetFileName(pdf), Nothing)
        Console.WriteLine("  PDF 分类: " & cls.ToString())
        Console.WriteLine("  发票类型判断: " & If(parse.InvoiceTypeText, "(未识别)"))
        Console.WriteLine(String.Format("  识别状态: {0}, 来源: {1}, 明细行数: {2}", res.Status, res.Source, res.Invoice.LineItems.Count))

        ' 候选评分前 3 名（命名核心字段），便于定位择优是否合理
        Console.WriteLine("  —— 候选评分前3名（命名核心字段）——")
        For Each fld As String In New String() {InvoiceFieldNames.InvoiceNumber, InvoiceFieldNames.InvoiceDate, InvoiceFieldNames.SellerName, InvoiceFieldNames.TotalWithTax}
            PrintTop3Candidates(parse, fld)
        Next

        If opts.DumpLayout Then
            Console.WriteLine("  —— 逐页行（页/行/坐标/文本）——")
            Dim li As Integer = 0
            For Each ln As PdfTextLine In extract.Lines
                Console.WriteLine(String.Format("    P{0} L{1,-3} Y={2,7:N1} X={3,6:N1}  {4}", ln.PageIndex, li, ln.Y, ln.Words(0).X, ln.Text))
                li += 1
            Next
        End If

        If opts.DumpCandidates Then
            Console.WriteLine("  —— 字段候选（含评分）——")
            For Each c As GeneralInvoiceFieldCandidate In parse.Candidates
                Console.WriteLine("    " & c.ToString())
            Next
            Console.WriteLine("  —— 择优结果 ——")
            For Each kv As KeyValuePair(Of String, GeneralInvoiceFieldCandidate) In parse.Chosen
                Console.WriteLine(String.Format("    {0} = {1}  (分={2:F1})", kv.Key, kv.Value.Value, kv.Value.ConfidenceScore))
            Next
        End If

        Dim inv As InvoiceInfo = res.Invoice
        PrintField("发票号码", inv.InvoiceNumber)
        PrintField("发票代码", inv.InvoiceCode)
        PrintField("开票日期", inv.InvoiceDate)
        PrintField("购买方名称", inv.BuyerName)
        PrintField("销售方名称", inv.SellerName)
        PrintField("价税合计", inv.TotalWithTax)
        PrintField("税率", inv.TaxRate)
        PrintField("税额", inv.TaxAmount)

        If opts.DumpLineItems OrElse inv.LineItems.Count > 0 Then
            Console.WriteLine("  —— 商品明细 ——")
            For Each it As InvoiceLineItem In inv.LineItems
                Console.WriteLine(String.Format("    #{0} 名称={1} 规格={2} 数量={3} 单价={4} 金额={5} 税率={6} 税额={7}",
                                                it.LineIndex, Dash(it.ItemName), Dash(it.Specification), Dash(it.Quantity),
                                                Dash(it.UnitPrice), Dash(it.Amount), Dash(it.TaxRate), Dash(it.TaxAmount)))
            Next
        End If

        For Each note As String In parse.Notes
            Console.WriteLine("  提示: " & note)
        Next

        Dim missing As List(Of String) = KeyFieldEvaluator.GetMissingGeneralInvoiceNamingFields(inv)
        Console.WriteLine(String.Format("  本地字段完整度: {0}/3（开票日期/金额/销售方名称）", 3 - missing.Count))
        Console.WriteLine("  命名核心缺失: " & If(missing.Count = 0, "无", String.Join("、", missing.ToArray())))
        Console.WriteLine("  是否建议回退在线 OCR: " & (missing.Count > 0))

        Dim plan As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv, InvoiceDocumentType.VatInvoice, Path.GetFileName(pdf),
                                                                        "20260101000000", Nothing, 1, res.Source.ToString(), NamingTemplates.DefaultGeneralInvoiceTemplate)
        Console.WriteLine("  常规发票命名预览: " & plan.FileName & If(plan.FallbackTriggered, " (fallback)", ""))

        ' 回归对比
        If expected IsNot Nothing Then
            Dim key As String = Path.GetFileName(pdf).ToLowerInvariant()
            Dim exp As Dictionary(Of String, String) = Nothing
            If expected.TryGetValue(key, exp) Then
                Return CompareExpected(inv, exp)
            Else
                Console.WriteLine("  （期望文件中无此 PDF 的条目，跳过对比）")
            End If
        End If
        Return New Integer() {0, 0}
    End Function

    ''' <summary>打印某字段评分前 3 名候选（诊断择优是否合理）。</summary>
    Private Sub PrintTop3Candidates(parse As GeneralInvoiceParseResult, fieldName As String)
        Dim list As New List(Of GeneralInvoiceFieldCandidate)()
        For Each c As GeneralInvoiceFieldCandidate In parse.Candidates
            If c.FieldName = fieldName Then list.Add(c)
        Next
        If list.Count = 0 Then
            Console.WriteLine("    " & fieldName & ": (无候选)")
            Return
        End If
        list.Sort(Function(a, b) b.ConfidenceScore.CompareTo(a.ConfidenceScore))
        Dim sb As New StringBuilder()
        sb.Append("    ").Append(fieldName).Append(": ")
        Dim n As Integer = 0
        For Each c As GeneralInvoiceFieldCandidate In list
            n += 1
            If n > 3 Then Exit For
            If n > 1 Then sb.Append(" | ")
            sb.Append(String.Format("{0}(分{1:F1},{2})", c.Value, c.ConfidenceScore, c.SourceRule))
        Next
        Console.WriteLine(sb.ToString())
    End Sub

    ''' <summary>与期望字段对比，打印每个字段的通过/失败与实际/期望，返回 {通过,失败}。</summary>
    Private Function CompareExpected(inv As InvoiceInfo, exp As Dictionary(Of String, String)) As Integer()
        Dim pass As Integer = 0, fail As Integer = 0
        Console.WriteLine("  —— 回归对比 ——")
        pass += CompareOne("开票日期", NormDate(inv.InvoiceDate), NormDate(GetExp(exp, "date")), fail)
        pass += CompareOne("金额", LocalTextNormalizer.NormalizeAmount(inv.TotalWithTax), LocalTextNormalizer.NormalizeAmount(GetExp(exp, "amount")), fail)
        pass += CompareOne("销售方名称", inv.SellerName, GetExp(exp, "seller"), fail)
        If Not String.IsNullOrEmpty(GetExp(exp, "number")) Then
            pass += CompareOne("发票号码", inv.InvoiceNumber, GetExp(exp, "number"), fail)
        End If
        If Not String.IsNullOrEmpty(GetExp(exp, "items")) Then
            pass += CompareOne("明细行数", inv.LineItems.Count.ToString(), GetExp(exp, "items"), fail)
        End If
        ' 重新数失败
        Dim total As Integer = 3
        If Not String.IsNullOrEmpty(GetExp(exp, "number")) Then total += 1
        If Not String.IsNullOrEmpty(GetExp(exp, "items")) Then total += 1
        fail = total - pass
        Return New Integer() {pass, fail}
    End Function

    Private Function CompareOne(name As String, actual As String, expectedVal As String, ByRef failDummy As Integer) As Integer
        Dim a As String = If(actual, "").Trim()
        Dim e As String = If(expectedVal, "").Trim()
        If String.IsNullOrEmpty(e) Then
            Console.WriteLine(String.Format("    - {0}: 跳过（无期望值）", name))
            Return 1 ' 无期望值视为不计失败
        End If
        If String.Equals(a, e, StringComparison.Ordinal) Then
            Console.WriteLine(String.Format("    PASS {0}: {1}", name, a))
            Return 1
        End If
        Console.WriteLine(String.Format("    FAIL {0}: 实际='{1}' 期望='{2}'", name, a, e))
        Return 0
    End Function

    Private Function NormDate(s As String) As String
        Dim ymd As String = LocalTextNormalizer.NormalizeDateToYmd(s)
        Return If(ymd, If(s, ""))
    End Function

    Private Function GetExp(exp As Dictionary(Of String, String), key As String) As String
        Dim v As String = Nothing
        If exp IsNot Nothing AndAlso exp.TryGetValue(key, v) Then Return v
        Return Nothing
    End Function

    ''' <summary>读取期望文件（.json 数组 或 .csv）。key=PDF 文件名(小写)，value=规范化字段字典。</summary>
    Private Function LoadExpected(path As String) As Dictionary(Of String, Dictionary(Of String, String))
        Try
            If Not File.Exists(path) Then
                Console.WriteLine("期望文件不存在：" & path)
                Return Nothing
            End If
            Dim map As New Dictionary(Of String, Dictionary(Of String, String))()
            If path.EndsWith(".json", StringComparison.OrdinalIgnoreCase) Then
                Dim ser As New System.Web.Script.Serialization.JavaScriptSerializer()
                Dim obj As Object = ser.DeserializeObject(File.ReadAllText(path))
                Dim arr = TryCast(obj, Object())
                If arr IsNot Nothing Then
                    For Each row In arr
                        Dim d = TryCast(row, Dictionary(Of String, Object))
                        If d Is Nothing Then Continue For
                        Dim rec As Dictionary(Of String, String) = CanonRecord(d)
                        Dim pdfName As String = GetExp(rec, "pdf")
                        If Not String.IsNullOrEmpty(pdfName) Then map(pdfName.ToLowerInvariant()) = rec
                    Next
                End If
            Else
                Dim lines As String() = File.ReadAllLines(path)
                If lines.Length < 2 Then Return map
                Dim headers As String() = lines(0).Split(","c)
                For r As Integer = 1 To lines.Length - 1
                    If String.IsNullOrWhiteSpace(lines(r)) Then Continue For
                    Dim cols As String() = lines(r).Split(","c)
                    Dim d As New Dictionary(Of String, Object)()
                    For ci As Integer = 0 To Math.Min(headers.Length, cols.Length) - 1
                        d(headers(ci).Trim()) = cols(ci).Trim()
                    Next
                    Dim rec As Dictionary(Of String, String) = CanonRecord(d)
                    Dim pdfName As String = GetExp(rec, "pdf")
                    If Not String.IsNullOrEmpty(pdfName) Then map(pdfName.ToLowerInvariant()) = rec
                Next
            End If
            Return map
        Catch ex As Exception
            Console.WriteLine("读取期望文件失败：" & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>把期望记录列名规范化为 pdf/date/amount/seller/number/items（兼容中英文列名）。</summary>
    Private Function CanonRecord(d As Dictionary(Of String, Object)) As Dictionary(Of String, String)
        Dim rec As New Dictionary(Of String, String)()
        For Each kv As KeyValuePair(Of String, Object) In d
            Dim k As String = kv.Key.Trim().ToLowerInvariant()
            Dim v As String = If(kv.Value Is Nothing, "", kv.Value.ToString())
            Select Case k
                Case "pdf", "file", "filename", "文件名", "文件" : rec("pdf") = v
                Case "date", "开票日期", "期望开票日期" : rec("date") = v
                Case "amount", "金额", "期望金额" : rec("amount") = v
                Case "seller", "销售方名称", "期望销售方名称", "销售方" : rec("seller") = v
                Case "number", "发票号码", "期望发票号码" : rec("number") = v
                Case "items", "明细行数", "期望明细行数" : rec("items") = v
            End Select
        Next
        Return rec
    End Function

    ''' <summary>合并多个 PDF 为一个（--merge &lt;out.pdf&gt; &lt;pdf1&gt; &lt;pdf2&gt; ...）。</summary>
    Private Sub RunMerge(positionals As List(Of String))
        If positionals Is Nothing OrElse positionals.Count < 3 Then
            Console.WriteLine("用法: OfflineTester.exe --merge <输出.pdf> <pdf1> <pdf2> [pdf3 ...]")
            Environment.ExitCode = 2
            Return
        End If
        Dim outPath As String = positionals(0)
        Dim inputs As New List(Of String)()
        For i As Integer = 1 To positionals.Count - 1
            inputs.Add(positionals(i))
        Next
        Dim svc As New PdfMergeService()
        Dim r As Result = svc.Merge(inputs, outPath)
        If r.IsSuccess Then
            Console.WriteLine("合并成功: " & outPath)
            Try
                Using doc = UglyToad.PdfPig.PdfDocument.Open(outPath)
                    Console.WriteLine("  合并后页数: " & doc.NumberOfPages)
                End Using
            Catch ex As Exception
                Console.WriteLine("  警告: 合并文件无法打开: " & ex.Message)
                Environment.ExitCode = 1
            End Try
        Else
            Console.WriteLine("合并失败: " & r.Message)
            Environment.ExitCode = 1
        End If
    End Sub

    ''' <summary>解析一份（脱敏）百度 OCR JSON 文件，跑通 解析器 + 映射器，打印结果。不联网。</summary>
    Private Sub ParseOcrJsonFile(jsonPath As String)
        Try
            If Not File.Exists(jsonPath) Then
                Console.WriteLine("文件不存在: " & jsonPath)
                Environment.ExitCode = 2
                Return
            End If
            Dim json As String = File.ReadAllText(jsonPath)
            Console.WriteLine("解析 OCR JSON: " & Path.GetFileName(jsonPath))
            PrintMappedFromJson(json)
        Catch ex As Exception
            Console.WriteLine("解析失败: " & ExceptionFormatter.ToUserMessage(ex))
            Environment.ExitCode = 1
        End Try
    End Sub

    ''' <summary>把一段百度 multiple_invoice JSON 跑通 解析器+映射器 并打印。</summary>
    Friend Sub PrintMappedFromJson(json As String)
        Dim parser As New BaiduMultipleInvoiceResponseParser()
        Dim doc As BaiduParsedDocument = parser.Parse(json)
        Console.WriteLine("  票据数 words_result_num=" & doc.WordsResultNum & ", 解析出 items=" & doc.Items.Count &
                          If(String.IsNullOrEmpty(doc.ParseError), "", ", 解析错误=" & doc.ParseError))

        Dim result As New InvoiceRecognitionResult()
        Dim mapper As New BaiduInvoiceFieldMapper()
        Dim mapped As Boolean = mapper.Map(doc, result)
        Console.WriteLine("  映射: " & mapped & ", 类型=" & result.DocumentType & ", 字段数=" & result.Fields.Count)

        Dim inv As InvoiceInfo = result.Invoice
        PrintField("发票号码", inv.InvoiceNumber)
        PrintField("开票日期", inv.InvoiceDate)
        PrintField("销售方名称", inv.SellerName)
        PrintField("价税合计", inv.TotalWithTax)
        Console.WriteLine("  行程明细数量: " & If(inv.Trips IsNot Nothing, inv.Trips.Count, 0))
        If inv.Trips IsNot Nothing Then
            Dim i As Integer = 0
            For Each t As InvoiceTripInfo In inv.Trips
                i += 1
                Console.WriteLine(String.Format("    行程#{0}: 车型={1}, 起点={2}, 终点={3}, 金额={4}",
                                                i, Dash(t.ServiceType), Dash(t.StartLocation), Dash(t.EndLocation), Dash(t.TripAmount)))
            Next
        End If
        If inv.ExtendedFields IsNot Nothing AndAlso inv.ExtendedFields.Count > 0 Then
            Console.WriteLine("  扩展字段数（未知/未强类型化）: " & inv.ExtendedFields.Count)
        End If
    End Sub

    Private Function ResolveFiles(target As String) As List(Of String)
        Dim files As New List(Of String)()
        If Directory.Exists(target) Then
            files.AddRange(Directory.GetFiles(target, "*.pdf"))
        ElseIf File.Exists(target) Then
            files.Add(target)
        End If
        Return files
    End Function

    Private Sub PrintHelp()
        Console.WriteLine("iWorkHelper 离线测试工具 OfflineTester")
        Console.WriteLine("")
        Console.WriteLine("用法:")
        Console.WriteLine("  OfflineTester.exe <pdf文件或目录> [选项]")
        Console.WriteLine("")
        Console.WriteLine("选项:")
        Console.WriteLine("  --local-only            仅本地文本解析，不调用在线 OCR")
        Console.WriteLine("  --ocr                   本地优先，字段不足/图片型时回退在线 OCR")
        Console.WriteLine("  --force-ocr             强制调用在线 OCR（本地充分也调用）")
        Console.WriteLine("  --save-response <目录>  保存脱敏后的 OCR 原始响应到指定目录")
        Console.WriteLine("  --template <模板>       用自定义统一命名模板做命名预览（默认 {乘车日期}_{金额}_{出发地点}_{到达地点}）")
        Console.WriteLine("  --parse-ocr-json <文件> 解析一份(脱敏)百度 OCR JSON，跑通解析器+映射器（不联网）")
        Console.WriteLine("  --dump-local-debug      转储本地解析诊断到 local-debug\（含票据信息，已被 .gitignore 忽略，勿提交）")
        Console.WriteLine("  --classify              输出每个 PDF 的分流分类（滴滴/常规发票/未识别）及命名")
        Console.WriteLine("  --preflight [目录]      归档前预检查（目录/OCR/模板/临时/日志）")
        Console.WriteLine("  --general-invoice       常规发票本地识别诊断（分类/类型/候选/择优/明细/缺失/命名预览）")
        Console.WriteLine("  --dump-layout           配合 --general-invoice：打印逐页行（页/行/坐标/文本）")
        Console.WriteLine("  --dump-candidates       配合 --general-invoice：打印字段候选与评分")
        Console.WriteLine("  --dump-line-items       配合 --general-invoice：打印商品明细")
        Console.WriteLine("  --compare-expected <文件> 常规发票回归对比（json/csv 期望；不提交真实样例/期望）")
        Console.WriteLine("  --validate-config [目录] 同 --preflight，校验配置有效性")
        Console.WriteLine("  --simulate-error <码名>  展示某错误码的用户友好文案（list 查看全部）")
        Console.WriteLine("  --selftest              运行内置自测：DPAPI 往返 / 命名模板边界 / 合成 OCR JSON")
        Console.WriteLine("  --save-baidu-config     用环境变量 AK/SK 写入本机加密配置（SK 经 DPAPI 加密）")
        Console.WriteLine("  --use-config            使用本机加密配置文件（DPAPI 解密 SK）执行 OCR")
        Console.WriteLine("  --help, -h              显示本帮助")
        Console.WriteLine("")
        Console.WriteLine("在线 OCR 需设置环境变量（不写入代码/日志）:")
        Console.WriteLine("  set BAIDU_OCR_AK=你的APIKey")
        Console.WriteLine("  set BAIDU_OCR_SK=你的SecretKey")
        Console.WriteLine("")
        Console.WriteLine("示例:")
        Console.WriteLine("  OfflineTester.exe sample")
        Console.WriteLine("  OfflineTester.exe sample\a.pdf --local-only")
        Console.WriteLine("  OfflineTester.exe sample --force-ocr --save-response ocr-raw")
        Console.WriteLine("")
        Console.WriteLine("命名模板支持的变量（与设置页一致，用 {变量名}）:")
        Console.WriteLine("  " & String.Join(" ", ArchiveNamingRule.SupportedPlaceholders().Select(Function(s) "{" & s & "}").ToArray()))
    End Sub

End Module

''' <summary>命令行选项解析结果。</summary>
Friend Class CliOptions
    Public Property TargetPath As String
    Public Property UseOcr As Boolean
    Public Property ForceOcr As Boolean
    Public Property LocalOnly As Boolean
    Public Property SaveResponseDir As String
    Public Property InvoiceTemplateOverride As String
    Public Property ParseOcrJsonPath As String
    Public Property SelfTest As Boolean
    Public Property SaveBaiduConfig As Boolean
    Public Property UseConfig As Boolean
    Public Property MergeMode As Boolean
    Public Property Positionals As New List(Of String)()
    Public Property DumpLocalDebug As Boolean
    Public Property LocalDebugDir As String
    Public Property Classify As Boolean
    Public Property ValidateConfig As Boolean
    Public Property Preflight As Boolean
    Public Property GeneralInvoice As Boolean
    Public Property DumpLayout As Boolean
    Public Property DumpCandidates As Boolean
    Public Property DumpLineItems As Boolean
    Public Property CompareExpectedPath As String
    Public Property SimulateError As String
    Public Property ShowHelp As Boolean
    Public Property ErrorMessage As String

    Public Shared Function Parse(args As String()) As CliOptions
        Dim o As New CliOptions()
        If args Is Nothing OrElse args.Length = 0 Then
            o.ShowHelp = True
            Return o
        End If
        Dim needArgHint As String = " 需要一个参数"

        Dim i As Integer = 0
        While i < args.Length
            Dim a As String = args(i)
            Select Case a.ToLowerInvariant()
                Case "--help", "-h", "/?"
                    o.ShowHelp = True
                Case "--ocr"
                    o.UseOcr = True
                Case "--force-ocr"
                    o.ForceOcr = True
                Case "--local-only"
                    o.LocalOnly = True
                Case "--save-response"
                    If i + 1 < args.Length Then
                        o.SaveResponseDir = args(i + 1)
                        i += 1
                    Else
                        o.ErrorMessage = "--save-response 需要一个目录参数"
                    End If
                Case "--template"
                    If i + 1 < args.Length Then
                        o.InvoiceTemplateOverride = args(i + 1)
                        i += 1
                    Else
                        o.ErrorMessage = "--template" & needArgHint
                    End If
                Case "--parse-ocr-json"
                    If i + 1 < args.Length Then
                        o.ParseOcrJsonPath = args(i + 1)
                        i += 1
                    Else
                        o.ErrorMessage = "--parse-ocr-json" & needArgHint
                    End If
                Case "--selftest", "--self-test"
                    o.SelfTest = True
                Case "--save-baidu-config"
                    o.SaveBaiduConfig = True
                Case "--use-config"
                    o.UseConfig = True
                Case "--merge"
                    o.MergeMode = True
                Case "--dump-local-debug"
                    o.DumpLocalDebug = True
                    o.LocalDebugDir = "local-debug"
                Case "--classify"
                    o.Classify = True
                Case "--validate-config"
                    o.ValidateConfig = True
                Case "--preflight"
                    o.Preflight = True
                Case "--general-invoice", "--general"
                    o.GeneralInvoice = True
                Case "--dump-layout"
                    o.DumpLayout = True
                Case "--dump-candidates"
                    o.DumpCandidates = True
                Case "--dump-line-items"
                    o.DumpLineItems = True
                Case "--compare-expected"
                    If i + 1 < args.Length Then
                        o.CompareExpectedPath = args(i + 1)
                        i += 1
                    Else
                        o.ErrorMessage = "--compare-expected 需要一个 json/csv 文件参数"
                    End If
                Case "--simulate-error"
                    If i + 1 < args.Length Then
                        o.SimulateError = args(i + 1)
                        i += 1
                    Else
                        o.SimulateError = "list"
                    End If
                Case Else
                    If a.StartsWith("-", StringComparison.Ordinal) Then
                        o.ErrorMessage = "未知选项: " & a
                    Else
                        o.Positionals.Add(a)
                        If String.IsNullOrEmpty(o.TargetPath) Then o.TargetPath = a
                    End If
            End Select
            i += 1
        End While
        Return o
    End Function
End Class

''' <summary>
''' 内置自测：不联网、不依赖 Outlook，验证可在本机确定验证的代码路径：
'''  - DPAPI Secret 加解密往返（当前 Windows 用户，与 VSTO 同一 DPAPI 上下文）；
'''  - 命名模板边界（空字段折叠 / 未知占位符 / 非法字符 / 超长）；
'''  - 合成（伪造、非敏感）百度 multiple_invoice JSON 的 解析器 + 映射器。
''' 注意：合成 JSON 仅验证代码路径，不代表真实百度字段名已校准。
''' </summary>
Friend Module SelfTestRunner

    Private _pass As Integer = 0
    Private _fail As Integer = 0

    Public Sub Run()
        Console.WriteLine("========== OfflineTester 自测 ==========")
        TestDpapi()
        TestNaming()
        TestSyntheticOcrJson()
        TestRecognitionMerge()
        TestErrorHandling()
        TestClassification()
        TestArchiveRunGuard()
        TestGeneralInvoiceLocal()
        TestLayoutClustering()
        Console.WriteLine("")
        Console.WriteLine(String.Format("自测结果：通过 {0}，失败 {1}", _pass, _fail))
        If _fail > 0 Then Environment.ExitCode = 1
    End Sub

    Private Sub TestDpapi()
        Console.WriteLine("[1] DPAPI 加解密往返")
        Try
            Dim sample As String = "iWorkHelper-DPAPI-selftest-not-a-real-key"
            Dim enc As String = SecretProtector.Protect(sample)
            Check("加密带 DPAPI: 前缀", SecretProtector.IsProtected(enc))
            Check("密文不等于明文", enc <> sample)
            Dim ok As Boolean
            Dim dec As String = SecretProtector.TryUnprotect(enc, ok)
            Check("解密成功标志", ok)
            Check("解密还原一致", dec = sample)
            ' 明文遗留值应原样返回（迁移路径）
            Dim ok2 As Boolean
            Dim passthrough As String = SecretProtector.TryUnprotect("legacy-plaintext", ok2)
            Check("明文遗留原样返回", ok2 AndAlso passthrough = "legacy-plaintext")
            Console.WriteLine("    （注：此为当前 Windows 用户的真实 DPAPI 往返；Outlook 会话内的保存/加载/迁移仍需人工验证）")
        Catch ex As Exception
            Fail("DPAPI 异常: " & ex.Message)
        End Try
    End Sub

    Private Sub TestNaming()
        Console.WriteLine("[2] 统一命名模板（默认 {乘车日期}_{金额}_{出发地点}_{到达地点}）")
        Try
            ' 默认模板 + 业务变量优先级：乘车日期=行程出发日期>开票日期；金额=行程金额>价税合计
            Dim inv As New InvoiceInfo()
            inv.InvoiceDate = "2026年05月26日"
            inv.TotalWithTax = "138.46"
            inv.SetExtendedField("行程起止日期", "2026-05-18")
            inv.Trips.Add(New InvoiceTripInfo With {.RowIndex = 1, .DepartureTime = "2026-05-18", .StartLocation = "上海虹桥站", .EndLocation = "上海市徐汇区", .TripAmount = "138.46"})
            Dim p0 As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv, InvoiceDocumentType.VatInvoice, "a.pdf", "20260101000000")
            Console.WriteLine("    默认模板命名: " & p0.FileName)
            Check("乘车日期优先行程出发日期(20260518)", p0.FileName.StartsWith("20260518_", StringComparison.Ordinal))
            Check("金额取行程金额(138.46)", p0.FileName.Contains("138.46"))
            Check("含出发地点(起点)", p0.FileName.Contains("上海虹桥站"))
            Check("含到达地点(终点)", p0.FileName.Contains("上海市徐汇区"))
            Check("规则名=统一模板", p0.RuleName = "统一模板")

            ' 空字段折叠：起点为空，不应出现连续下划线
            Dim inv1 As New InvoiceInfo()
            inv1.SetExtendedField("行程起止日期", "2026-05-18")
            inv1.Trips.Add(New InvoiceTripInfo With {.RowIndex = 1, .DepartureTime = "2026-05-18", .StartLocation = "", .EndLocation = "上海市徐汇区", .TripAmount = "138.46"})
            Dim p1 As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv1, InvoiceDocumentType.RideTripStatement, "a.pdf", "t")
            Check("空字段不产生连续下划线", Not p1.FileName.Contains("__"))
            Console.WriteLine("    命名(空起点): " & p1.FileName)

            ' 未知占位符
            Dim t2 As New NamingTemplates() With {.UnifiedTemplate = "{乘车日期}_{不存在}_{金额}"}
            Dim p2 As ArchiveNamePlan = New ArchiveNamingRule(t2).BuildPlan(inv, InvoiceDocumentType.VatInvoice, "a.pdf", "t")
            Check("未知占位符被记录", p2.UnknownPlaceholders IsNot Nothing AndAlso p2.UnknownPlaceholders.Contains("不存在"))

            ' 非法字符清理
            Dim inv3 As New InvoiceInfo()
            inv3.SetExtendedField("行程起止日期", "2026-01-01")
            inv3.Trips.Add(New InvoiceTripInfo With {.DepartureTime = "2026-01-01", .StartLocation = "a/b:c*d?e", .EndLocation = "x", .TripAmount = "1.00"})
            Dim p3 As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv3, InvoiceDocumentType.RideTripStatement, "a.pdf", "t")
            Check("非法字符被清理", (New String() {"/"c, ":"c, "*"c, "?"c}).All(Function(c) Not p3.FileName.Contains(c)))

            ' 字段严重不足 → fallback（未识别票据_邮件主题_时间戳）
            Dim emptyInv As New InvoiceInfo()
            Dim p4 As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(emptyInv, InvoiceDocumentType.Unknown, "att.pdf", "20260101", "示例邮件主题", 2)
            Check("字段严重不足触发 fallback", p4.FallbackTriggered)
            Check("fallback 含‘未识别票据’", p4.FileName.Contains("未识别票据"))
            Check("fallback 用邮件主题", p4.FileName.Contains("示例邮件主题"))
            Console.WriteLine("    fallback 命名: " & p4.FileName)

            ' 标点模板：占位符全空但渲染结果非空白（如「（）」）→ 仍须回退，不得产出垃圾文件名（O-14）
            Dim t5 As New NamingTemplates() With {.UnifiedTemplate = "{金额}（{出发地点}）"}
            Dim p5 As ArchiveNamePlan = New ArchiveNamingRule(t5).BuildPlan(emptyInv, InvoiceDocumentType.Unknown, "att.pdf", "20260101", "示例邮件主题", 2)
            Check("标点模板字段全空仍触发 fallback", p5.FallbackTriggered)
            Check("标点模板不产出垃圾文件名", Not p5.FileName.Contains("（）"))
            Console.WriteLine("    命名(标点模板/字段全空): " & p5.FileName)
        Catch ex As Exception
            Fail("命名异常: " & ex.Message)
        End Try
    End Sub

    Private Sub TestSyntheticOcrJson()
        Console.WriteLine("[3] 合成百度 multiple_invoice JSON（伪造数据，仅验证代码路径）")
        Try
            ' 用单引号书写再替换为双引号，避免大量转义。数据为伪造、非敏感。
            Dim json As String = ("{" &
                "'words_result_num':3," &
                "'pdf_file_size':'2'," &
                "'words_result':[" &
                " {'type':'vat_invoice','result':{" &
                "   'InvoiceNum':[{'word':'26127000000268665697','probability':{'average':0.99}}]," &
                "   'InvoiceDate':[{'word':'2026年05月26日'}]," &
                "   'SellerName':[{'word':'滴滴出行科技有限公司'}]," &
                "   'AmountInFiguers':[{'word':'138.46'}]," &
                "   'CommodityTaxRate':[{'word':'3%'}]," &
                "   'SomeUnknownField':[{'word':'X1'}]" &
                " }}," &
                " {'type':'taxi_online_ticket','result':{" &
                "   'Time':[{'word':'05-18 19:31'}]," &
                "   'Start':[{'word':'A地'}]," &
                "   'Destination':[{'word':'B地'}]," &
                "   'TotalFare':[{'word':'138.46'}]" &
                " }}," &
                " {'type':'foobar_unknown_type','result':{'AnyField':[{'word':'Y1'}]}}" &
                "]}").Replace("'"c, """"c)

            Dim parser As New BaiduMultipleInvoiceResponseParser()
            Dim doc As BaiduParsedDocument = parser.Parse(json)
            Check("解析出 3 张票据", doc.Items.Count = 3)

            ' 类型映射：未知 type 不失败
            Check("未知 type 归 Other", BaiduInvoiceTypeMapper.Map("foobar_unknown_type") = InvoiceDocumentType.Other)
            Check("vat_invoice 映射正确", BaiduInvoiceTypeMapper.Map("vat_invoice") = InvoiceDocumentType.VatInvoice)
            Check("taxi_online_ticket 映射正确", BaiduInvoiceTypeMapper.Map("taxi_online_ticket") = InvoiceDocumentType.RideTripStatement)

            Dim result As New InvoiceRecognitionResult()
            Dim mapper As New BaiduInvoiceFieldMapper()
            Dim mapped As Boolean = mapper.Map(doc, result)
            Check("映射返回 true", mapped)
            Check("发票号码映射", result.Invoice.InvoiceNumber = "26127000000268665697")
            Check("销售方映射", result.Invoice.SellerName = "滴滴出行科技有限公司")
            Check("价税合计映射", result.Invoice.TotalWithTax = "138.46")
            Check("未知字段进扩展字段", result.Invoice.ExtendedFields.ContainsKey("SomeUnknownField"))
            Check("行程被解析(>=1条)", result.Invoice.Trips IsNot Nothing AndAlso result.Invoice.Trips.Count >= 1)
            Check("置信度被保留", result.Fields.Any(Function(f) f.Confidence > 0 AndAlso f.Confidence < 1.0000001 AndAlso f.Confidence >= 0.9))
            Console.WriteLine("    映射类型=" & result.DocumentType & ", 字段数=" & result.Fields.Count & ", 行程数=" & result.Invoice.Trips.Count)
        Catch ex As Exception
            Fail("合成 JSON 异常: " & ex.Message)
        End Try
    End Sub

    Private Sub TestRecognitionMerge()
        Console.WriteLine("[4] 识别结果合并（发票结果 + 行程单结果 -> 一个）")
        Try
            ' 发票-only 结果
            Dim invR As New InvoiceRecognitionResult()
            invR.Source = RecognitionSource.LocalText
            invR.DocumentType = InvoiceDocumentType.VatInvoice
            invR.Status = RecognitionStatus.Success
            invR.Invoice.InvoiceNumber = "26127000000268665697"
            invR.Invoice.InvoiceDate = "2026年05月26日"
            invR.Invoice.SellerName = "滴滴出行科技有限公司"
            invR.Invoice.TotalWithTax = "138.46"

            ' 行程单-only 结果
            Dim tripR As New InvoiceRecognitionResult()
            tripR.Source = RecognitionSource.BaiduOcr
            tripR.DocumentType = InvoiceDocumentType.RideTripStatement
            tripR.Status = RecognitionStatus.PartialSuccess
            tripR.Invoice.Trips.Add(New InvoiceTripInfo With {.RowIndex = 1, .Passenger = "示例乘客", .StartLocation = "A地", .EndLocation = "B地", .TripAmount = "138.46"})

            Dim merger As New InvoiceRecognitionMerger()
            Check("发票 PDF 分类为 Invoice", merger.Classify(invR) = InvoiceRecognitionMerger.PdfKind.Invoice)
            Check("行程单 PDF 分类为 Trip", merger.Classify(tripR) = InvoiceRecognitionMerger.PdfKind.Trip)

            Dim merged As InvoiceRecognitionResult = merger.Merge(New List(Of InvoiceRecognitionResult) From {tripR, invR})
            Check("合并后含发票号码", merged.Invoice.InvoiceNumber = "26127000000268665697")
            Check("合并后含行程明细", merged.Invoice.Trips IsNot Nothing AndAlso merged.Invoice.Trips.Count >= 1)
            Check("合并后类型=发票", merged.DocumentType = InvoiceDocumentType.VatInvoice)
            Check("合并来源=Mixed（本地+在线）", merged.Source = RecognitionSource.Mixed)

            ' 合并后用统一模板命名（发票字段仍保留在结果中）
            Dim plan As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(merged.Invoice, merged.DocumentType, "invoice.pdf", "t", "邮件主题示例", 2, merged.Source.ToString())
            Check("合并后命名用统一模板", plan.RuleName = "统一模板")
            Check("合并后命名非空且非 fallback", Not plan.FallbackTriggered AndAlso Not String.IsNullOrEmpty(plan.FileName))
            Console.WriteLine("    合并命名预览: " & plan.FileName)
        Catch ex As Exception
            Fail("识别合并异常: " & ex.Message)
        End Try
    End Sub

    Private Sub TestErrorHandling()
        Console.WriteLine("[5] 容错：错误分类 / 预检查 / 敏感信息不泄漏")
        Try
            ' 每个错误码都有非空用户提示
            Dim allHaveMsg As Boolean = True
            For Each code As AppErrorCode In [Enum].GetValues(GetType(AppErrorCode))
                If code = AppErrorCode.None Then Continue For
                Dim e As AppError = UserFriendlyMessageProvider.Describe(code)
                If String.IsNullOrWhiteSpace(e.UserMessage) Then allHaveMsg = False
            Next
            Check("所有错误码有用户提示", allHaveMsg)

            ' OCR 失败分类
            Check("401 → 密钥错误", UserFriendlyMessageProvider.ClassifyOcrFailure("Client authentication failed", 401) = AppErrorCode.OcrAuthFailed)
            Check("timeout → 超时", UserFriendlyMessageProvider.ClassifyOcrFailure("The operation has timed out", 0) = AppErrorCode.OcrTimeout)
            Check("quota → 额度不足", UserFriendlyMessageProvider.ClassifyOcrFailure("quota limit", 403) = AppErrorCode.OcrUnauthorizedOrQuota)

            ' 预检查：未配置归档目录 → 阻断
            Dim c As New ArchivePreflightChecker()
            Dim r1 As ArchivePreflightResult = c.Check(1, Nothing, New BaiduOcrOptions(), "{乘车日期}_{金额}")
            Check("未配置归档目录→阻断", r1.HasBlocking)
            ' 预检查：未选中邮件 → 阻断
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "iwh_pre_" & Guid.NewGuid().ToString("N").Substring(0, 8))
            Directory.CreateDirectory(tempDir)
            Dim r2 As ArchivePreflightResult = c.Check(0, tempDir, New BaiduOcrOptions(), "{乘车日期}_{金额}")
            Check("未选中邮件→阻断", r2.HasBlocking)
            ' 预检查：有效目录 + 有选中 → 无阻断
            Dim r3 As ArchivePreflightResult = c.Check(1, tempDir, New BaiduOcrOptions(), "{乘车日期}_{金额}")
            Check("有效目录+已选中→不阻断", Not r3.HasBlocking)
            Try : Directory.Delete(tempDir, True) : Catch : End Try

            ' 友好文案不含堆栈/英文异常关键字
            Dim unk As AppError = UserFriendlyMessageProvider.Describe(AppErrorCode.Unknown)
            Check("未知异常文案友好(无 stack/null ref)", Not unk.UserMessage.ToLowerInvariant().Contains("exception") AndAlso Not unk.UserMessage.Contains("null"))
        Catch ex As Exception
            Fail("容错测试异常: " & ex.Message)
        End Try
    End Sub

    Private Sub TestClassification()
        Console.WriteLine("[6] 分流分类：滴滴/常规发票/未识别 + 命名")
        Try
            Dim c As New MailProcessingClassifier()

            ' 滴滴发票 PDF（含发票号 + 行程）
            Dim didiInv As New InvoiceRecognitionResult()
            didiInv.Invoice.InvoiceNumber = "26127000000268665697"
            didiInv.Invoice.SellerName = "滴滴出行科技有限公司"
            didiInv.Invoice.Trips.Add(New InvoiceTripInfo With {.StartLocation = "A", .EndLocation = "B", .TripAmount = "138.46"})
            didiInv.DocumentType = InvoiceDocumentType.RideTripStatement
            Check("滴滴发票PDF→DidiInvoicePdf", c.ClassifyPdf(didiInv, "滴滴出行科技有限公司 发票号码", "a.pdf", "滴滴") = PdfAttachmentClassification.DidiInvoicePdf)

            ' 滴滴行程单 PDF（无发票号 + 行程单文本）
            Dim didiTrip As New InvoiceRecognitionResult()
            didiTrip.DocumentType = InvoiceDocumentType.RideTripStatement
            Check("滴滴行程单PDF→DidiTripPdf", c.ClassifyPdf(didiTrip, "滴滴出行-行程单 TRIP TABLE 上车时间 共1笔行程", "t.pdf", "行程单") = PdfAttachmentClassification.DidiTripPdf)

            ' 常规发票 PDF（发票号 + 无滴滴/行程）
            Dim gen As New InvoiceRecognitionResult()
            gen.Invoice.InvoiceNumber = "11111111111111111111"
            gen.Invoice.SellerName = "某某科技有限公司"
            gen.Invoice.InvoiceDate = "2026年05月26日"
            gen.Invoice.TotalWithTax = "138.46"
            gen.DocumentType = InvoiceDocumentType.VatInvoice
            Check("常规发票PDF→GeneralInvoicePdf", c.ClassifyPdf(gen, "增值税电子普通发票 发票代码 发票号码 购买方 销售方 价税合计", "b.pdf", "发票") = PdfAttachmentClassification.GeneralInvoicePdf)

            ' 未识别 PDF（无任何特征）
            Dim unk As New InvoiceRecognitionResult()
            Check("未识别PDF→UnknownPdf", c.ClassifyPdf(unk, "some random content", "x.pdf", "附件") = PdfAttachmentClassification.UnknownPdf)

            ' 邮件类型
            Check("含滴滴→DidiInvoiceMail", c.ClassifyMail(New List(Of PdfAttachmentClassification) From {PdfAttachmentClassification.DidiInvoicePdf, PdfAttachmentClassification.DidiTripPdf}) = MailProcessingType.DidiInvoiceMail)
            Check("常规+未知→MixedPdfMail", c.ClassifyMail(New List(Of PdfAttachmentClassification) From {PdfAttachmentClassification.GeneralInvoicePdf, PdfAttachmentClassification.UnknownPdf}) = MailProcessingType.MixedPdfMail)
            Check("全常规→GeneralInvoiceMail", c.ClassifyMail(New List(Of PdfAttachmentClassification) From {PdfAttachmentClassification.GeneralInvoicePdf}) = MailProcessingType.GeneralInvoiceMail)
            Check("全未知→UnknownPdfOnlyMail", c.ClassifyMail(New List(Of PdfAttachmentClassification) From {PdfAttachmentClassification.UnknownPdf}) = MailProcessingType.UnknownPdfOnlyMail)
            Check("无PDF→NoPdfMail", c.ClassifyMail(New List(Of PdfAttachmentClassification)()) = MailProcessingType.NoPdfMail)

            ' 常规发票默认命名 {开票日期}_{金额}_{销售方名称}
            Dim gp As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(gen.Invoice, InvoiceDocumentType.VatInvoice, "b.pdf", "t", Nothing, 1, "LocalText", NamingTemplates.DefaultGeneralInvoiceTemplate)
            Check("常规发票命名=开票日期_金额_销售方", gp.FileName = "20260526_138.46_某某科技有限公司.pdf")
            Console.WriteLine("    常规发票命名: " & gp.FileName)

            ' 未识别命名 未识别_原始文件名
            Dim un As String = UnknownPdfNamingRule.BuildFileName("某某扫描件.pdf")
            Check("未识别命名=未识别_原名", un = "未识别_某某扫描件.pdf")
            Console.WriteLine("    未识别命名: " & un)
        Catch ex As Exception
            Fail("分类测试异常: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 归档运行锁 ArchiveRunGuard 自测：覆盖“防重复点击 / 释放后可再次获取 / 预检查失败与业务异常必释放”。
    ''' 这是对“点击归档立即提示已有任务运行”缺陷的回归测试，可脱离 Outlook 环境确定验证。
    ''' </summary>
    Private Sub TestArchiveRunGuard()
        Console.WriteLine("[7] 归档运行锁 ArchiveRunGuard（原子获取/释放/防重入）")
        Try
            Check("初始状态未运行", Not ArchiveRunGuard.IsRunning)

            Dim t1 As ArchiveRunToken = ArchiveRunGuard.TryAcquire("selftest-1")
            Check("首次可获取运行锁", t1 IsNot Nothing)
            Check("获取后标记为运行中", ArchiveRunGuard.IsRunning)

            Dim t2 As ArchiveRunToken = ArchiveRunGuard.TryAcquire("selftest-2")
            Check("运行中第二次获取失败（防重复点击/防并发）", t2 Is Nothing)

            t1.Dispose()
            Check("释放后回到未运行", Not ArchiveRunGuard.IsRunning)

            Dim t3 As ArchiveRunToken = ArchiveRunGuard.TryAcquire("selftest-3")
            Check("释放后可再次获取", t3 IsNot Nothing)
            t3.Dispose()
            t3.Dispose() ' 幂等
            Check("重复 Dispose 安全且已释放", Not ArchiveRunGuard.IsRunning)

            ' 模拟：预检查失败早退（Using 结束即释放）
            Using tk As ArchiveRunToken = ArchiveRunGuard.TryAcquire("selftest-preflight-fail")
                Check("Using 内已获取运行锁", tk IsNot Nothing)
            End Using
            Check("预检查失败早退后运行锁已释放", Not ArchiveRunGuard.IsRunning)

            ' 模拟：业务流程抛异常，Using 仍保证释放
            Dim releasedAfterEx As Boolean = False
            Try
                Using tk As ArchiveRunToken = ArchiveRunGuard.TryAcquire("selftest-exception")
                    Throw New InvalidOperationException("模拟归档业务异常（自测）")
                End Using
            Catch
                releasedAfterEx = Not ArchiveRunGuard.IsRunning
            End Try
            Check("业务异常后运行锁已释放", releasedAfterEx)

            Check("自测结束后为未运行（不残留）", Not ArchiveRunGuard.IsRunning)
        Catch ex As Exception
            Fail("运行锁测试异常: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 常规发票本地识别专项自测（合成、非敏感文本，仅验证代码路径）：
    ''' 委派、购/销不混淆、金额取价税合计而非明细金额、发票号码不误取、商品明细、字段不足回退 OCR、滴滴不受影响。
    ''' </summary>
    Private Sub TestGeneralInvoiceLocal()
        Console.WriteLine("[8] 常规发票本地识别（委派/候选评分/明细/OCR 回退）")
        Try
            ' —— 用例1：完整常规发票（数电票版式，购/销分区 + 明细 + 小写金额）——
            Dim full As String = String.Join(vbLf, New String() {
                "电子发票（普通发票）",
                "发票号码：25317000000012345678",
                "开票日期：2026-05-20",
                "购买方信息",
                "名称：上海买方科技有限公司",
                "统一社会信用代码/纳税人识别号：91310000MA1FL0000X",
                "销售方信息",
                "名称：北京卖方咨询有限公司",
                "统一社会信用代码/纳税人识别号：91110000MA00ABCD1Y",
                "项目名称 规格型号 单位 数量 单价 金额 税率 税额",
                "*信息技术服务*软件开发服务 100.00 6% 6.00",
                "合计 ¥100.00 ¥6.00",
                "价税合计（大写）壹佰零陆圆整（小写）¥106.00"})

            Dim r As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "general.pdf", .ExtractedText = full})
            Dim inv As InvoiceInfo = r.Invoice

            Check("常规发票委派到专用识别器", r.RecognizerName = "GeneralInvoiceLocalRecognizer")
            Check("类型=VatInvoice", r.DocumentType = InvoiceDocumentType.VatInvoice)
            Check("状态=Success", r.Status = RecognitionStatus.Success)
            Check("发票号码=20位数电号", inv.InvoiceNumber = "25317000000012345678")
            Check("开票日期=2026-05-20", inv.InvoiceDate = "2026-05-20")
            Check("销售方=北京卖方（不被购买方覆盖）", inv.SellerName = "北京卖方咨询有限公司")
            Check("购买方=上海买方", inv.BuyerName = "上海买方科技有限公司")
            Check("购销不混淆", inv.SellerName <> inv.BuyerName)
            Check("金额取价税合计106.00（非明细100.00）", inv.TotalWithTax = "106.00")
            Check("商品明细>=1条", inv.LineItems IsNot Nothing AndAlso inv.LineItems.Count >= 1)
            If inv.LineItems.Count >= 1 Then
                Check("明细金额=100.00", inv.LineItems(0).Amount = "100.00")
                Check("明细税率=6%", inv.LineItems(0).TaxRate = "6%")
            End If
            ' 命名预览：{开票日期}_{金额}_{销售方名称}
            Dim plan As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(inv, InvoiceDocumentType.VatInvoice, "general.pdf",
                                                                            "t", Nothing, 1, "LocalText", NamingTemplates.DefaultGeneralInvoiceTemplate)
            Check("常规发票命名=开票日期_金额_销售方", plan.FileName = "20260520_106.00_北京卖方咨询有限公司.pdf")
            Console.WriteLine("    常规发票命名: " & plan.FileName)

            ' —— 用例2：缺销售方 → 部分成功 + 判定字段不足（应回退在线 OCR）——
            Dim noSeller As String = String.Join(vbLf, New String() {
                "增值税电子普通发票",
                "发票号码：25317000000099998888",
                "开票日期：2026年06月01日",
                "价税合计（小写）¥50.00"})
            Dim r2 As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "noseller.pdf", .ExtractedText = noSeller})
            Check("缺销售方=部分成功", r2.Status = RecognitionStatus.PartialSuccess)
            Check("缺销售方→命名不充分（触发 OCR 回退）", Not KeyFieldEvaluator.IsNamingSufficient(r2.Invoice, InvoiceDocumentType.VatInvoice))
            Check("缺失字段含销售方名称", KeyFieldEvaluator.GetMissingGeneralInvoiceNamingFields(r2.Invoice).Contains("销售方名称"))

            ' —— 用例3：滴滴发票不受影响（仍由原识别器处理，不委派）——
            Dim didi As String = String.Join(vbLf, New String() {
                "滴滴出行科技有限公司",
                "发票号码：11111111111111111111",
                "滴滴出行-行程单",
                "共1笔行程"})
            Dim r3 As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "didi.pdf", .ExtractedText = didi})
            Check("滴滴类型=RideTripStatement（未委派）", r3.DocumentType = InvoiceDocumentType.RideTripStatement)
            Check("滴滴仍由原 LocalTextInvoiceRecognizer 处理", r3.RecognizerName = "LocalTextInvoiceRecognizer")

            ' —— 用例4：数电票（无发票代码）+ 号码与税号/校验码共存 + 多条明细 ——
            Dim shudian As String = String.Join(vbLf, New String() {
                "电子发票（普通发票）",
                "发票号码：25317000000012345678",
                "开票日期：2026-07-05",
                "购买方名称：甲科技有限公司",
                "统一社会信用代码/纳税人识别号：91310000MA1FL0000X",
                "销售方名称：乙咨询有限公司",
                "统一社会信用代码/纳税人识别号：91110000MA00ABCD1Y",
                "项目名称 规格型号 单位 数量 单价 金额 税率 税额",
                "*运输服务*客运服务 200.00 3% 6.00",
                "*住宿服务*住宿费 100.00 6% 6.00",
                "价税合计（小写）¥312.00",
                "校验码：98765432109876543210"})
            Dim r4 As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "shudian.pdf", .ExtractedText = shudian})
            Check("数电票发票号码正确（不误取税号/校验码）", r4.Invoice.InvoiceNumber = "25317000000012345678")
            Check("数电票发票代码为空", String.IsNullOrEmpty(r4.Invoice.InvoiceCode))
            Check("数电票销售方=乙咨询", r4.Invoice.SellerName = "乙咨询有限公司")
            Check("数电票购买方=甲科技", r4.Invoice.BuyerName = "甲科技有限公司")
            Check("数电票金额=312.00（价税合计非明细）", r4.Invoice.TotalWithTax = "312.00")
            Check("多条商品明细=2", r4.Invoice.LineItems.Count = 2)

            ' —— 用例5：商品名称换行（金额在续行）——
            Dim wrap As String = String.Join(vbLf, New String() {
                "增值税电子普通发票",
                "发票号码：25317000000055556666",
                "开票日期：2026年08月10日",
                "销售方名称：丙软件有限公司",
                "项目名称 金额 税率 税额",
                "*信息技术服务*软件开发及",
                "技术咨询服务 500.00 6% 30.00",
                "价税合计（小写）¥530.00"})
            Dim r5 As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "wrap.pdf", .ExtractedText = wrap})
            Check("名称换行合并为1条明细", r5.Invoice.LineItems.Count = 1)
            If r5.Invoice.LineItems.Count = 1 Then
                Check("换行名称已拼接", r5.Invoice.LineItems(0).ItemName.Contains("软件开发") AndAlso r5.Invoice.LineItems(0).ItemName.Contains("技术咨询"))
                Check("续行金额补齐=500.00", r5.Invoice.LineItems(0).Amount = "500.00")
            End If
            Check("换行样例金额=价税合计530.00", r5.Invoice.TotalWithTax = "530.00")

            ' —— 用例6：仅销售方名称/备注提及“滴滴、行程单、网约车”→ 不得误判为行程单（O-38）——
            Dim notRide As String = String.Join(vbLf, New String() {
                "增值税电子普通发票",
                "发票号码：25317000000077778888",
                "开票日期：2026年09月01日",
                "销售方名称：滴滴出行科技有限公司",
                "价税合计（小写）¥66.00",
                "备注：行程单见附件（网约车报销）"})
            Dim r6 As InvoiceRecognitionResult = New LocalTextInvoiceRecognizer().Recognize(
                New RecognitionContext With {.OriginalFileName = "notride.pdf", .ExtractedText = notRide})
            Check("仅备注/销售方提及滴滴→不误判行程单", r6.DocumentType = InvoiceDocumentType.VatInvoice)
            Check("该发票仍解析出销售方", Not String.IsNullOrEmpty(r6.Invoice.SellerName))
            Console.WriteLine("    非行程单样例类型=" & r6.DocumentType & "，销售方=" &
                              If(String.IsNullOrWhiteSpace(r6.Invoice.SellerName), "-", r6.Invoice.SellerName))
        Catch ex As Exception
            Fail("常规发票识别测试异常: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' [9] 坐标行聚类不变量（O-40）：行内 Y 跨度永远不超过容差，不存在链式漂移。
    ''' </summary>
    Private Sub TestLayoutClustering()
        Console.WriteLine("[9] 坐标行聚类（行内 Y 跨度 <= 容差，不链式合并）")
        Try
            ' 9 单位间距（远大于容差 4）必须落到不同的行
            Dim spaced As New List(Of PdfTextWord)()
            For i As Integer = 0 To 2
                spaced.Add(New PdfTextWord With {.Text = "W" & i, .X = i * 10, .Right = i * 10 + 5, .Y = 100 - i * 9})
            Next
            Dim linesSpaced As List(Of PdfTextLine) = PdfTextLayoutExtractor.ClusterIntoLines(spaced, 1)
            Check("9 单位间距不合并（3 行）", linesSpaced.Count = 3)

            ' 3 单位间距的连续链：任何一行的 Y 跨度都不得超过容差（超过即为链式漂移）
            Dim chain As New List(Of PdfTextWord)()
            For i As Integer = 0 To 5
                chain.Add(New PdfTextWord With {.Text = "C" & i, .X = i * 10, .Right = i * 10 + 5, .Y = 100 - i * 3})
            Next
            Dim linesChain As List(Of PdfTextLine) = PdfTextLayoutExtractor.ClusterIntoLines(chain, 1)
            Dim maxSpread As Double = 0
            For Each ln As PdfTextLine In linesChain
                For Each w As PdfTextWord In ln.Words
                    Dim d As Double = Math.Abs(ln.Y - w.Y)
                    If d > maxSpread Then maxSpread = d
                Next
            Next
            Check("链式词不漂移（行内跨度<=容差）", maxSpread <= PdfTextLayoutExtractor.LineYTolerance)
            Check("链式词被拆成多行（未整体合并）", linesChain.Count >= 2)
            Console.WriteLine("    9 单位间距行数=" & linesSpaced.Count & "；3 单位链行数=" & linesChain.Count &
                              "；最大行内跨度=" & maxSpread.ToString(System.Globalization.CultureInfo.InvariantCulture))
        Catch ex As Exception
            Fail("行聚类测试异常: " & ex.Message)
        End Try
    End Sub

    Private Sub Check(name As String, ok As Boolean)
        If ok Then
            _pass += 1
            Console.WriteLine("    PASS - " & name)
        Else
            _fail += 1
            Console.WriteLine("    FAIL - " & name)
        End If
    End Sub

    Private Sub Fail(msg As String)
        _fail += 1
        Console.WriteLine("    FAIL - " & msg)
    End Sub

End Module
