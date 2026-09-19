Imports System.IO
Imports System.Xml.Linq

''' <summary>
''' OCR 配置读取器。第三阶段起 **以 SettingsForm 写入的 My.Settings 为主要来源**。
''' 配置来源优先级：
'''  1) My.Settings（SettingsForm 设置界面写入）——主入口；
'''  2) 若 My.Settings 中未填 AK/SK，则尝试兼容旧的外部 XML 文件
'''     %AppData%\iWorkHelper\baidu-ocr.config.xml（历史遗留，可选）。
''' 任何异常都回退为禁用配置，不抛出。密钥不写入日志。
''' </summary>
Public Module OcrConfigProvider

    Public Const ConfigFileName As String = "baidu-ocr.config.xml"

    ''' <summary>
    ''' 读取百度 OCR 配置（主入口）。
    ''' </summary>
    Public Function LoadBaiduOptions() As BaiduOcrOptions
        Dim options As BaiduOcrOptions
        Try
            options = LoadFromSettings()
        Catch ex As Exception
            AppLogger.Error("从 My.Settings 读取 OCR 配置失败，尝试回退 XML。", ex)
            options = New BaiduOcrOptions()
        End Try

        ' 若设置中未提供 AK/SK，则以外部加密配置文件（BaiduXmlConfigStore）作为完整来源。
        ' 该文件的 Secret Key 为 DPAPI 加密，读取时解密。用于"开发期配置落地"。
        If String.IsNullOrWhiteSpace(options.ApiKey) OrElse String.IsNullOrWhiteSpace(options.SecretKey) Then
            Try
                Dim fromXml As BaiduOcrOptions = BaiduXmlConfigStore.Load()
                If fromXml IsNot Nothing Then
                    options = fromXml ' 完整采用外部配置（含 Enabled/URL/开关等）
                End If
            Catch ex As Exception
                AppLogger.Error("读取外部 OCR 配置失败（忽略）。", ex)
            End Try
        End If

        ' —— 端点白名单 ——
        ' 外部配置文件（%AppData%\iWorkHelper\baidu-ocr.config.xml）当前用户可写，且可能随
        ' 漫游/同步 profile 落到别处。若不加校验，能写该文件的一方即可把 TokenUrl 指向自己的
        ' 主机，随后 BaiduAccessTokenProvider 会向该主机 POST client_secret。
        ' 因此非受信任域的端点一律回退为官方默认值（并在日志中说明）。
        Const DefaultTokenUrl As String = "https://aip.baidubce.com/oauth/2.0/token"
        Const DefaultOcrUrl As String = "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice"

        If Not IsTrustedEndpoint(options.TokenUrl) Then
            AppLogger.Warn("Token 端点不在受信任域内，已回退为默认百度端点（如需自建网关请加入白名单常量）。")
            options.TokenUrl = DefaultTokenUrl
        End If
        If Not IsTrustedEndpoint(options.OcrApiUrl) Then
            AppLogger.Warn("OCR 端点不在受信任域内，已回退为默认百度端点（如需自建网关请加入白名单常量）。")
            options.OcrApiUrl = DefaultOcrUrl
        End If

        ' —— 统一收口（单一控制入口）——
        ' 在线 OCR 是否可调用，仅由「编译期是否允许在线解析（BuildFeatures.OnlineParserEnabled）」
        ' 与「当前 ParseMode 是否为在线解析」共同决定；不再依赖已废弃的 OcrEnabled 复选框。
        ' 内网版（未定义 INTERNET_BUILD）恒为 False。
        options.Enabled = IsOnlineParseAllowed()

        AppLogger.Info("OCR 配置已加载：" & options.ToSafeSummary())
        Return options
    End Function

    ''' <summary>
    ''' 受信任的端点主机后缀（可在此追加企业自建网关域名）。
    ''' 匹配规则：主机等于后缀，或以 ".后缀" 结尾（避免 evil-baidubce.com 之类伪装）。
    ''' </summary>
    Private ReadOnly TrustedHostSuffixes As String() = {"baidubce.com"}

    ''' <summary>
    ''' 校验端点 URL 是否可安全承载密钥：必须是 https，且主机落在受信任域内。
    ''' </summary>
    Public Function IsTrustedEndpoint(url As String) As Boolean
        If String.IsNullOrWhiteSpace(url) Then Return False
        Try
            Dim u As New Uri(url.Trim())
            If Not String.Equals(u.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then Return False
            Dim host As String = u.Host
            For Each suffix As String In TrustedHostSuffixes
                If String.Equals(host, suffix, StringComparison.OrdinalIgnoreCase) OrElse
                   host.EndsWith("." & suffix, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 当前是否允许调用在线 OCR = 编译期允许在线解析 且 ParseMode 为在线解析。
    ''' 内网版本恒为 False。读取 My.Settings 失败按“本地”处理。
    ''' </summary>
    Public Function IsOnlineParseAllowed() As Boolean
        If Not BuildFeatures.OnlineParserEnabled Then Return False
        Dim mode As String = Nothing
        Try
            mode = My.Settings.ParseMode
        Catch
        End Try
        Return String.Equals(If(mode, "Local").Trim(), "Online", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' 从 My.Settings 构建配置对象。
    ''' </summary>
    Private Function LoadFromSettings() As BaiduOcrOptions
        Dim s = My.Settings
        Dim o As New BaiduOcrOptions()
        ' 注意：o.Enabled 不再取自旧的 OcrEnabled 复选框字段；
        ' 其最终值在 LoadBaiduOptions 中由 IsOnlineParseAllowed() 统一收口决定。
        o.ApiKey = If(s.BaiduApiKey, String.Empty).Trim()
        ' Secret Key 经 DPAPI 加密存储，此处解密使用（解密失败返回空，不阻断）。
        o.SecretKey = If(ProtectedSettingsProvider.GetSecretKey(), String.Empty).Trim()
        o.OcrApiUrl = If(String.IsNullOrWhiteSpace(s.BaiduOcrApiUrl), "https://aip.baidubce.com/rest/2.0/ocr/v1/multiple_invoice", s.BaiduOcrApiUrl.Trim())
        o.TokenUrl = If(String.IsNullOrWhiteSpace(s.BaiduTokenUrl), "https://aip.baidubce.com/oauth/2.0/token", s.BaiduTokenUrl.Trim())
        o.TimeoutMilliseconds = If(s.OcrTimeoutMs > 0, s.OcrTimeoutMs, 30000)
        o.ReturnProbability = s.OcrReturnProbability
        o.ReturnLocation = s.OcrReturnLocation
        o.VerifyParameter = s.OcrVerifyParameter
        o.MaxPages = If(s.OcrMaxPages > 0, s.OcrMaxPages, 1)
        o.PreferLocalParse = s.PreferLocalParse
        o.AutoFallbackToOcr = s.AutoFallbackToOcr
        Return o
    End Function

    Public Function GetConfigFilePath() As String
        Return BaiduXmlConfigStore.GetPath()
    End Function

End Module
