Imports System.Windows.Forms
Imports System.Reflection
Imports System.Deployment.Application

Public Class SettingsForm

    Private isLoading As Boolean

    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isLoading = True
        Try
            ' 应用内外网版本的布局
            ApplyEditionLayout()

            ' 设置窗口打开时，检查并执行明文密钥迁移（如果需要）。
            ' 这样避免在 Startup 时做不必要的 My.Settings.Save()。
            Try
                Dim migrated = ProtectedSettingsProvider.MigratePlaintextIfNeeded()
                If migrated Then
                    AppLogger.Info("设置窗口打开时，检测到明文 Secret Key 已迁移为 DPAPI 加密。")
                End If
            Catch ex As Exception
                AppLogger.Warn("明文密钥迁移失败（不影响设置窗口打开）：" & ex.Message)
            End Try

            txtFolderPath.Text = My.Settings.ArchiveFolderPath

            ' 隐藏在线解析区域（内网版本）已在 ApplyEditionLayout() 中处理
            grpOcr.Visible = BuildFeatures.OnlineParserEnabled
            rdoOnlineParse.Visible = BuildFeatures.OnlineParserEnabled

            ' 配置回退：内网版本强制使用本地解析
            Dim parseMode As String = My.Settings.ParseMode
            If Not BuildFeatures.OnlineParserEnabled Then
                parseMode = "Local"  ' 内网版本强制本地解析
            End If

            If String.Equals(parseMode, "Online", StringComparison.OrdinalIgnoreCase) Then
                rdoOnlineParse.Checked = True
            Else
                rdoLocalParse.Checked = True
            End If

            ' 载入 OCR 配置：使用“有效配置”（My.Settings 优先，未填时取本机加密配置文件）。
            ' 注：是否启用在线 OCR 已由“本地解析/在线解析”单选项唯一控制，不再有独立复选框。
            Dim eff As BaiduOcrOptions = OcrConfigProvider.LoadBaiduOptions()
            txtApiKey.Text = If(eff.ApiKey, String.Empty)
            txtSecretKey.Text = If(eff.SecretKey, String.Empty)
            ' 若 My.Settings 中已有 SK 但无法在当前 Windows 用户解密，提示重输（不阻断）。
            Dim decryptFailed As Boolean
            ProtectedSettingsProvider.GetSecretKey(decryptFailed)
            If decryptFailed AndAlso String.IsNullOrEmpty(eff.SecretKey) Then
                MessageBox.Show("已保存的 Secret Key 无法在当前 Windows 用户下解密（可能来自其它用户/机器）。请重新输入并保存。",
                                "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            txtApiUrl.Text = If(String.IsNullOrWhiteSpace(eff.OcrApiUrl), "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice", eff.OcrApiUrl)
            txtTokenUrl.Text = If(String.IsNullOrWhiteSpace(eff.TokenUrl), "https://aip.baidubce.com/oauth/2.0/token", eff.TokenUrl)
            numTimeout.Value = ClampDecimal(eff.TimeoutMilliseconds, numTimeout.Minimum, numTimeout.Maximum, 30000)
            numMaxPages.Value = ClampDecimal(eff.MaxPages, numMaxPages.Minimum, numMaxPages.Maximum, 1)
            chkProbability.Checked = eff.ReturnProbability
            chkLocation.Checked = eff.ReturnLocation
            chkVerify.Checked = eff.VerifyParameter
            chkPreferLocal.Checked = eff.PreferLocalParse
            chkAutoFallback.Checked = eff.AutoFallbackToOcr

            ' 命名模板：滴滴发票（txtInvoiceTemplate） + 常规发票（txtTripTemplate） + 未识别（txtUnknownTemplate）。
            txtInvoiceTemplate.Text = If(String.IsNullOrEmpty(My.Settings.UnifiedNameTemplate), NamingTemplates.DefaultUnifiedTemplate, My.Settings.UnifiedNameTemplate)
            txtTripTemplate.Text = If(String.IsNullOrEmpty(My.Settings.GeneralInvoiceNameTemplate), NamingTemplates.DefaultGeneralInvoiceTemplate, My.Settings.GeneralInvoiceNameTemplate)
            txtUnknownTemplate.Text = If(String.IsNullOrEmpty(My.Settings.UnknownNameTemplate), NamingTemplates.DefaultUnknownTemplate, My.Settings.UnknownNameTemplate)

            ' 版本显示：包含版本号和版本类型（内网版本会显示"（内网版）"）
            Dim versionText As String = "版本：" & GetApplicationVersion()
            If Not String.IsNullOrEmpty(BuildFeatures.EditionDisplay) Then
                versionText &= BuildFeatures.EditionDisplay
            End If
            lblVersion.Text = versionText
        Finally
            isLoading = False
        End Try
    End Sub

    Private Sub btnBrowseFolder_Click(sender As Object, e As EventArgs) Handles btnBrowseFolder.Click
        folderBrowserDialog1.SelectedPath = txtFolderPath.Text

        If folderBrowserDialog1.ShowDialog(Me) = DialogResult.OK Then
            txtFolderPath.Text = folderBrowserDialog1.SelectedPath
            SaveFolderPath(txtFolderPath.Text)
        End If
    End Sub

    Private Sub txtFolderPath_Leave(sender As Object, e As EventArgs) Handles txtFolderPath.Leave
        SaveFolderPath(txtFolderPath.Text)
    End Sub

    Private Sub rdoLocalParse_CheckedChanged(sender As Object, e As EventArgs) Handles rdoLocalParse.CheckedChanged
        If isLoading OrElse Not rdoLocalParse.Checked Then
            Return
        End If
        My.Settings.ParseMode = "Local"
        My.Settings.Save()
    End Sub

    Private Sub rdoOnlineParse_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOnlineParse.CheckedChanged
        If isLoading OrElse Not rdoOnlineParse.Checked Then
            Return
        End If
        ' 内网版本不允许选择在线解析
        If Not BuildFeatures.OnlineParserEnabled Then
            rdoLocalParse.Checked = True
            Return
        End If
        My.Settings.ParseMode = "Online"
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' 保存全部设置（含 OCR）。带基本校验：启用在线 OCR 但 AK/SK 未填写时提示且不保存。
    ''' 注意：不在任何日志/弹窗中输出完整 AK/SK。
    ''' </summary>
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim apiKey As String = txtApiKey.Text.Trim()
        Dim secretKey As String = txtSecretKey.Text.Trim()
        Dim apiUrl As String = txtApiUrl.Text.Trim()
        Dim tokenUrl As String = txtTokenUrl.Text.Trim()

        ' —— 归档目录校验 ——
        Dim folder As String = If(txtFolderPath.Text, "").Trim()
        If Not String.IsNullOrEmpty(folder) Then
            Dim invalidPath As Boolean = False
            Try
                System.IO.Path.GetFullPath(folder)
            Catch
                invalidPath = True
            End Try
            If invalidPath Then
                MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.ArchiveFolderPathInvalid).ToUserText(),
                                "归档目录", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If Not System.IO.Directory.Exists(folder) Then
                If MessageBox.Show("归档目录不存在，是否创建？" & Environment.NewLine & folder,
                                   "归档目录", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    If Not PathHelper.EnsureDirectory(folder) Then
                        MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.ArchiveFolderNotWritable).ToUserText(),
                                        "归档目录", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                Else
                    Return ' 用户不创建则不保存无效目录
                End If
            End If
        End If

        ' —— 在线 OCR 校验 ——
        ' 校验入口由“在线解析”单选项决定（外网版且选择在线解析时才校验 AK/SK/URL 等）。
        Dim onlineParseSelected As Boolean = BuildFeatures.OnlineParserEnabled AndAlso rdoOnlineParse.Checked
        If onlineParseSelected Then
            If String.IsNullOrEmpty(apiKey) OrElse String.IsNullOrEmpty(secretKey) Then
                MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.OcrConfigMissing).ToUserText(),
                                "配置不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrEmpty(apiUrl) OrElse Not IsValidHttpUrl(apiUrl) Then
                MessageBox.Show("接口地址无效，请填写正确的 http(s) 接口地址。",
                                "配置不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If String.IsNullOrEmpty(tokenUrl) OrElse Not IsValidHttpUrl(tokenUrl) Then
                MessageBox.Show("Token 地址无效，请填写正确的 http(s) 地址。",
                                "配置不完整", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If numTimeout.Value < 3000 OrElse numTimeout.Value > 120000 Then
                MessageBox.Show("超时时间需在 3000 到 120000 毫秒之间。",
                                "配置不合理", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If numMaxPages.Value < 1 Then
                MessageBox.Show("最大识别页数需为正整数。",
                                "配置不合理", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        ' 命名模板未知变量提示（允许保存）
        WarnUnknownTemplateVars(txtInvoiceTemplate.Text.Trim())

        Try
            My.Settings.ArchiveFolderPath = If(txtFolderPath.Text, String.Empty).Trim()
            ' 旧的 OcrEnabled 字段仅为配置兼容而保留（业务代码不再依赖它）：
            ' 令其镜像“在线解析”单选项的有效值，避免旧配置读取到过期状态。
            My.Settings.OcrEnabled = onlineParseSelected
            My.Settings.BaiduApiKey = apiKey
            ' Secret Key 经 DPAPI 加密后保存（不明文存 user.config）。
            ' 加密不可用时中止本次保存，绝不明文落盘。
            If Not ProtectedSettingsProvider.SetSecretKey(secretKey) Then
                MessageBox.Show("无法加密保存 Secret Key（Windows DPAPI 不可用），本次设置未保存，以避免密钥以明文写入配置文件。",
                                "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            My.Settings.BaiduOcrApiUrl = If(String.IsNullOrEmpty(apiUrl), "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice", apiUrl)
            My.Settings.BaiduTokenUrl = If(String.IsNullOrEmpty(tokenUrl), "https://aip.baidubce.com/oauth/2.0/token", tokenUrl)
            My.Settings.OcrTimeoutMs = CInt(numTimeout.Value)
            My.Settings.OcrMaxPages = CInt(numMaxPages.Value)
            My.Settings.OcrReturnProbability = chkProbability.Checked
            My.Settings.OcrReturnLocation = chkLocation.Checked
            My.Settings.OcrVerifyParameter = chkVerify.Checked
            My.Settings.PreferLocalParse = chkPreferLocal.Checked
            My.Settings.AutoFallbackToOcr = chkAutoFallback.Checked
            ' 命名模板：滴滴 + 常规发票 + 未识别；留空则写入默认模板。
            Dim unifiedTpl As String = txtInvoiceTemplate.Text.Trim()
            If String.IsNullOrEmpty(unifiedTpl) Then unifiedTpl = NamingTemplates.DefaultUnifiedTemplate
            My.Settings.UnifiedNameTemplate = unifiedTpl
            Dim generalTpl As String = txtTripTemplate.Text.Trim()
            If String.IsNullOrEmpty(generalTpl) Then generalTpl = NamingTemplates.DefaultGeneralInvoiceTemplate
            My.Settings.GeneralInvoiceNameTemplate = generalTpl
            Dim unknownTpl As String = txtUnknownTemplate.Text.Trim()
            If String.IsNullOrEmpty(unknownTpl) Then unknownTpl = NamingTemplates.DefaultUnknownTemplate
            My.Settings.UnknownNameTemplate = unknownTpl
            My.Settings.Save()

            ' 保存后校验：确认存储的确实是 DPAPI 密文，再宣布成功。
            If Not ProtectedSettingsProvider.VerifyPersisted() Then
                MessageBox.Show("Secret Key 保存后校验失败：配置中不是加密值，请重试或检查用户配置文件权限。",
                                "保存校验失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ' 密钥可能变更，使已缓存的 token 失效。
            BaiduAccessTokenProvider.InvalidateCache()

            AppLogger.Info("设置已保存，设置窗口关闭。")
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            ' 不输出密钥；日志留详情，弹窗给友好文案。
            Dim err As New AppError(AppErrorCode.ConfigSaveFailed, ExceptionFormatter.ToUserMessage(ex))
            err.LogSelf()
            MessageBox.Show(err.ToUserText(), "错误", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    ''' <summary>命名模板含未知变量时提示（允许保存）。</summary>
    Private Sub WarnUnknownTemplateVars(template As String)
        If String.IsNullOrEmpty(template) Then Return
        Try
            Dim render As NamingRenderResult = NamingTemplateEngine.Render(template, New Dictionary(Of String, String)(), "x")
            ' 用空值渲染只为收集占位符；未知占位符 = 不在支持清单内的。
            Dim supported As New System.Collections.Generic.HashSet(Of String)(ArchiveNamingRule.SupportedPlaceholders())
            Dim unknown As New System.Collections.Generic.List(Of String)()
            For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(template, "\{([^{}]+)\}")
                Dim key As String = m.Groups(1).Value.Trim()
                If Not supported.Contains(key) AndAlso Not unknown.Contains(key) Then unknown.Add(key)
            Next
            If unknown.Count > 0 Then
                Dim msg As String = "命名模板中包含未知变量（不会被替换）：" & String.Join("、", unknown.ToArray())
                msg &= Environment.NewLine & "可点变量说明核对可用变量。设置仍会保存。"
                MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>“测试 OCR 配置”：仅测试 Token 获取（轻量连通性），不上传 PDF、不显示 token。</summary>
    Private Sub btnTestOcr_Click(sender As Object, e As EventArgs) Handles btnTestOcr.Click
        Dim ak As String = txtApiKey.Text.Trim()
        Dim sk As String = txtSecretKey.Text.Trim()
        If String.IsNullOrEmpty(ak) OrElse String.IsNullOrEmpty(sk) Then
            MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.OcrConfigMissing).ToUserText(),
                            "测试 OCR 配置", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim options As New BaiduOcrOptions With {
            .Enabled = True, .ApiKey = ak, .SecretKey = sk,
            .TokenUrl = If(String.IsNullOrEmpty(txtTokenUrl.Text.Trim()), "https://aip.baidubce.com/oauth/2.0/token", txtTokenUrl.Text.Trim()),
            .OcrApiUrl = If(String.IsNullOrEmpty(txtApiUrl.Text.Trim()), "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice", txtApiUrl.Text.Trim()),
            .TimeoutMilliseconds = CInt(numTimeout.Value)}

        Dim oldCursor = Me.Cursor
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim provider As New BaiduAccessTokenProvider()
            Dim tokenResult As BaiduAccessTokenResult = provider.GetToken(options)
            If tokenResult.Success Then
                AppLogger.Info("测试 OCR 配置：Token 获取成功（连通性正常）。")
                MessageBox.Show("百度 OCR 配置可用（连通性正常）。", "测试 OCR 配置", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Dim code As AppErrorCode = UserFriendlyMessageProvider.ClassifyOcrFailure(tokenResult.ErrorMessage, tokenResult.HttpStatusCode, tokenResult.BaiduError)
                Dim err As New AppError(code, "Token 获取失败（测试）")
                err.LogSelf()
                MessageBox.Show(err.ToUserText(), "测试 OCR 配置", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            AppLogger.Error("测试 OCR 配置异常。", ex)
            MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.OcrNetworkError).ToUserText(),
                            "测试 OCR 配置", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Me.Cursor = oldCursor
        End Try
    End Sub

    ''' <summary>展示统一命名模板支持的全部变量（单一来源：ArchiveNamingRule.SupportedPlaceholders）。</summary>
    Private Sub btnTplVars_Click(sender As Object, e As EventArgs) Handles btnTplVars.Click
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("最终归档命名规则支持以下变量（用 {变量名} 引用）：")
        sb.AppendLine()
        sb.AppendLine("【滴滴发票（默认 {乘车日期}_{金额}_{出发地点}_{到达地点}）】")
        sb.AppendLine("  {乘车日期} {金额} {出发地点} {到达地点}")
        sb.AppendLine("  取值：乘车日期=行程出发日期>开票日期；金额=行程金额>价税合计>金额；出发地点=起点，到达地点=终点")
        sb.AppendLine()
        sb.AppendLine("【常规发票（默认 {开票日期}_{金额}_{销售方名称}）】")
        sb.AppendLine("  {开票日期} {金额} {销售方名称} {发票号码} {价税合计} {原始文件名}")
        sb.AppendLine("  金额对常规发票=价税合计>金额")
        sb.AppendLine()
        sb.AppendLine("【未识别（默认 {原始文件名}）】")
        sb.AppendLine("  {原始文件名}")
        sb.AppendLine("  用于无法识别的 PDF 文件，默认保留源文件名")
        sb.AppendLine()
        sb.AppendLine("【通用】")
        sb.AppendLine("  {邮件主题} {原附件名} {原始文件名} {时间戳} {票据类型} {识别来源} {附件数量} {合并文件名}")
        sb.AppendLine()
        sb.AppendLine("【行程】")
        sb.AppendLine("  {出发日期} {出发时间} {到达时间} {乘车人} {起点} {终点}")
        sb.AppendLine("  {城市} {车型} {订单号} {里程} {行程金额}")
        sb.AppendLine()
        sb.AppendLine("【发票】")
        sb.AppendLine("  {开票日期} {销售方名称} {购买方名称} {发票号码} {发票代码}")
        sb.AppendLine("  {价税合计} {税额} {税率} {开票人} {项目名称}")
        sb.AppendLine()
        sb.AppendLine("说明：")
        sb.AppendLine("• {原始文件名} = 源 PDF 文件名（不含扩展名）；.pdf 扩展名由归档流程自动添加")
        sb.AppendLine("• 字段为空的变量会被自动跳过（不产生多余下划线）")
        sb.AppendLine("• 未知变量会被忽略并记录警告")
        sb.AppendLine("• 模板为空时自动回退为默认模板")
        sb.AppendLine("• 字段严重不足时用未识别 fallback（邮件主题_时间戳）")
        MessageBox.Show(sb.ToString(), "命名规则变量说明", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>用脱敏假数据预览最终归档文件名（单一模板，复用 ArchiveNamingRule / NamingTemplateEngine）。</summary>
    Private Sub btnTplPreview_Click(sender As Object, e As EventArgs) Handles btnTplPreview.Click
        Try
            Dim tpls As New NamingTemplates()
            tpls.UnifiedTemplate = txtInvoiceTemplate.Text.Trim()
            Dim rule As New ArchiveNamingRule(tpls)

            ' 脱敏假数据（非真实票据）：模拟合并后完整识别结果（发票 + 行程单）。
            Dim sample As New InvoiceInfo() With {
                .InvoiceDate = "2026年05月26日", .SellerName = "示例出行科技有限公司",
                .InvoiceNumber = "00000000000000000000", .TotalWithTax = "138.46", .Amount = "134.43"}
            sample.SetExtendedField("行程起止日期", "2026-05-18")
            sample.Trips.Add(New InvoiceTripInfo With {
                .RowIndex = 1, .DepartureTime = "2026-05-18", .StartLocation = "上海虹桥站",
                .EndLocation = "上海市徐汇区", .TripAmount = "138.46", .ServiceType = "快车"})

            Dim plan As ArchiveNamePlan = rule.BuildPlan(sample, InvoiceDocumentType.VatInvoice, "示例附件.pdf", "20260518090000", "示例邮件主题", 2, "Mixed")

            ' 常规发票预览（脱敏假数据）
            Dim genTpl As String = txtTripTemplate.Text.Trim()
            If String.IsNullOrEmpty(genTpl) Then genTpl = NamingTemplates.DefaultGeneralInvoiceTemplate
            Dim genSample As New InvoiceInfo() With {
                .InvoiceDate = "2026年05月26日", .SellerName = "某某科技有限公司", .InvoiceNumber = "11111111111111111111",
                .TotalWithTax = "138.46", .Amount = "134.43"}
            Dim genPlan As ArchiveNamePlan = New ArchiveNamingRule().BuildPlan(genSample, InvoiceDocumentType.VatInvoice, "发票示例.pdf", "20260526090000", Nothing, 1, "LocalText", genTpl)

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("最终归档文件名预览（脱敏假数据）：")
            sb.AppendLine()
            sb.AppendLine("滴滴发票：" & plan.FileName)
            AppendUnknown(sb, plan)
            sb.AppendLine("常规发票：" & genPlan.FileName)
            AppendUnknown(sb, genPlan)
            sb.AppendLine("未识别 PDF：" & UnknownPdfNamingRule.BuildFileName("某某扫描件.pdf"))
            If plan.FallbackTriggered Then
                sb.AppendLine()
                sb.AppendLine("提示：当前字段严重不足，已使用未识别 fallback 命名。")
            End If
            If String.IsNullOrWhiteSpace(txtInvoiceTemplate.Text) Then
                sb.AppendLine()
                sb.AppendLine("提示：模板为空，将使用默认模板 {乘车日期}_{金额}_{出发地点}_{到达地点}。")
            End If
            MessageBox.Show(sb.ToString(), "命名规则预览", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("预览失败：" & ExceptionFormatter.ToUserMessage(ex), "预览", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub AppendUnknown(sb As System.Text.StringBuilder, plan As ArchiveNamePlan)
        If plan IsNot Nothing AndAlso plan.UnknownPlaceholders IsNot Nothing AndAlso plan.UnknownPlaceholders.Count > 0 Then
            sb.AppendLine("  ⚠ 未知变量：" & String.Join("、", plan.UnknownPlaceholders.ToArray()))
        End If
    End Sub

    Private Sub SaveFolderPath(folderPath As String)
        If isLoading Then
            Return
        End If
        My.Settings.ArchiveFolderPath = If(folderPath, String.Empty).Trim()
        My.Settings.Save()
    End Sub

    Private Function IsValidHttpUrl(url As String) As Boolean
        Dim result As Uri = Nothing
        Return Uri.TryCreate(url, UriKind.Absolute, result) _
            AndAlso (result.Scheme = Uri.UriSchemeHttp OrElse result.Scheme = Uri.UriSchemeHttps)
    End Function

    Private Function ClampDecimal(value As Integer, min As Decimal, max As Decimal, fallback As Integer) As Decimal
        Dim v As Decimal = value
        If value <= 0 Then v = fallback
        If v < min Then v = min
        If v > max Then v = max
        Return v
    End Function

    Private Sub lblTplHint_Click(sender As Object, e As EventArgs)

    End Sub

    ''' <summary>获取用户展示版本，优先读取 InformationalVersion。</summary>
    Private Function GetApplicationVersion() As String
        Try
            Dim asm = Assembly.GetExecutingAssembly()
            Dim attr = CType(Attribute.GetCustomAttribute(asm, GetType(AssemblyInformationalVersionAttribute)), AssemblyInformationalVersionAttribute)
            If attr IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(attr.InformationalVersion) Then
                Return attr.InformationalVersion
            End If

            Dim version = asm.GetName().Version
            Return If(version Is Nothing, "未知", version.ToString())
        Catch
            Return "未知"
        End Try
    End Function

    ''' <summary>
    ''' 应用内外网版本的布局调整。
    ''' 设计器原始布局适配外网版本（包含在线 OCR）。
    ''' 内网版本隐藏解析模式选择和在线 OCR 配置，将命名规则分组框上移，缩短窗体高度。
    ''' </summary>
    Private Sub ApplyEditionLayout()
        If BuildFeatures.OnlineParserEnabled Then
            ' 外网版本：保持设计器原始布局
            Return
        End If

        ' 内网版本：应用紧凑布局
        Const GrpNamingHeight As Integer = 155
        Const GrpNamingNewY As Integer = 40
        Const LblVersionNewY As Integer = GrpNamingNewY + GrpNamingHeight + 8
        Const BtnSaveNewY As Integer = LblVersionNewY + 20
        Const FormNewHeight As Integer = BtnSaveNewY + 35

        ' 隐藏解析模式选择（内网版只能用本地）
        rdoLocalParse.Visible = False
        rdoOnlineParse.Visible = False

        ' 隐藏在线 OCR 分组框和相关说明
        grpOcr.Visible = False

        ' 上移命名规则分组框
        grpNaming.Location = New System.Drawing.Point(grpNaming.Location.X, GrpNamingNewY)

        ' 上移版本标签
        lblVersion.Location = New System.Drawing.Point(lblVersion.Location.X, LblVersionNewY)

        ' 上移保存按钮
        btnSave.Location = New System.Drawing.Point(btnSave.Location.X, BtnSaveNewY)

        ' 缩短窗体高度
        Me.ClientSize = New System.Drawing.Size(Me.ClientSize.Width, FormNewHeight)

        AppLogger.Debug("已应用内网版紧凑布局（隐藏在线解析、命名规则上移、窗体高度缩短）")
    End Sub

End Class
