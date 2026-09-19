Imports System.IO
Imports System.Globalization
Imports System.Xml
Imports System.Xml.Linq

''' <summary>
''' 百度 OCR 外部配置文件存储（%AppData%\iWorkHelper\baidu-ocr.config.xml）。
''' 特性：
'''  - **Secret Key 以 DPAPI（CurrentUser）加密保存**（带 DPAPI: 前缀），读时解密；
'''  - API Key/接口地址/开关等普通配置明文保存（非高度敏感）；
'''  - 不依赖 My.Settings，主程序与 OfflineTester 共用同一实现，避免两套规则；
'''  - 该文件属本机用户配置，已被 .gitignore 忽略，不得入库。
''' 用途：开发期把 OCR 配置安全落地到本机；主程序 OcrConfigProvider 在 My.Settings 未填时读取。
''' </summary>
Public Module BaiduXmlConfigStore

    Public Const ConfigFileName As String = "baidu-ocr.config.xml"

    Public Function GetPath() As String
        Return Path.Combine(PathHelper.GetAppDataRoot(), ConfigFileName)
    End Function

    Public Function Exists() As Boolean
        Try
            Return File.Exists(GetPath())
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 保存配置。plainSecret 为明文 Secret Key，写入前 DPAPI 加密。
    ''' 返回是否成功。**加密失败时中止保存**（绝不明文落盘）。
    ''' 采用"写临时文件 → 原子替换"，避免中断/满盘产生残破 XML。
    ''' </summary>
    Public Function Save(options As BaiduOcrOptions, plainSecret As String) As Boolean
        Dim tmp As String = Nothing
        Try
            If options Is Nothing Then options = New BaiduOcrOptions()

            Dim encOk As Boolean
            Dim encSecret As String = SecretProtector.TryProtect(If(plainSecret, String.Empty), encOk)
            If Not encOk Then
                AppLogger.Error("OCR 外部配置未保存：Secret Key 加密失败（已拒绝明文落盘）。")
                Return False
            End If
            If Not String.IsNullOrEmpty(plainSecret) AndAlso Not SecretProtector.IsProtected(encSecret) Then
                AppLogger.Error("OCR 外部配置未保存：加密结果不是 DPAPI 密文。")
                Return False
            End If

            PathHelper.EnsureDirectory(PathHelper.GetAppDataRoot())

            Dim doc As New XDocument(
                New XElement("BaiduOcr",
                    New XElement("Enabled", options.Enabled.ToString()),
                    New XElement("ApiKey", If(options.ApiKey, String.Empty)),
                    New XElement("SecretKey", If(encSecret, String.Empty)),
                    New XElement("TokenUrl", If(options.TokenUrl, String.Empty)),
                    New XElement("OcrApiUrl", If(options.OcrApiUrl, String.Empty)),
                    New XElement("TimeoutMilliseconds", options.TimeoutMilliseconds.ToString(CultureInfo.InvariantCulture)),
                    New XElement("ReturnProbability", options.ReturnProbability.ToString()),
                    New XElement("ReturnLocation", options.ReturnLocation.ToString()),
                    New XElement("VerifyParameter", options.VerifyParameter.ToString()),
                    New XElement("MaxPages", options.MaxPages.ToString(CultureInfo.InvariantCulture)),
                    New XElement("PreferLocalParse", options.PreferLocalParse.ToString()),
                    New XElement("AutoFallbackToOcr", options.AutoFallbackToOcr.ToString())))

            Dim target As String = GetPath()
            tmp = target & ".tmp"
            doc.Save(tmp)
            If File.Exists(target) Then
                File.Replace(tmp, target, Nothing)
            Else
                File.Move(tmp, target)
            End If
            tmp = Nothing
            AppLogger.Info("OCR 外部配置已保存（Secret Key 已 DPAPI 加密）。")
            Return True
        Catch ex As Exception
            AppLogger.Error("保存 OCR 外部配置失败。", ex)
            Return False
        Finally
            ' 失败时不要留下半成品临时文件。
            If Not String.IsNullOrEmpty(tmp) Then
                Try
                    If File.Exists(tmp) Then File.Delete(tmp)
                Catch
                End Try
            End If
        End Try
    End Function

    ''' <summary>
    ''' 读取配置。SecretKey 若为 DPAPI 密文则解密；文件不存在返回 Nothing。
    ''' 文件损坏与文件不存在都会返回 Nothing（调用方按"OCR 未配置"处理），但损坏会记日志。
    ''' 显式禁用 DTD 处理与外部解析器（XXE 加固；.NET 4.8 默认已安全，此处显式化）。
    ''' </summary>
    Public Function Load() As BaiduOcrOptions
        Try
            Dim p As String = GetPath()
            If Not File.Exists(p) Then
                Return Nothing
            End If

            Dim settings As New XmlReaderSettings()
            settings.DtdProcessing = DtdProcessing.Prohibit
            settings.XmlResolver = Nothing

            Dim root As XElement
            Using reader As XmlReader = XmlReader.Create(p, settings)
                Dim doc As XDocument = XDocument.Load(reader)
                root = doc.Root
            End Using

            If root Is Nothing Then
                AppLogger.Warn("OCR 外部配置为空文档：" & PrivacySafeFormatter.MaskPath(p))
                Return Nothing
            End If

            Dim o As New BaiduOcrOptions()
            o.Enabled = ParseBool(GetValue(root, "Enabled"), False)
            o.ApiKey = GetValue(root, "ApiKey")

            Dim storedSecret As String = GetValue(root, "SecretKey")
            o.SecretKey = If(SecretProtector.Unprotect(storedSecret), String.Empty)

            Dim apiUrl As String = GetValue(root, "OcrApiUrl")
            If Not String.IsNullOrWhiteSpace(apiUrl) Then o.OcrApiUrl = apiUrl
            Dim tokenUrl As String = GetValue(root, "TokenUrl")
            If Not String.IsNullOrWhiteSpace(tokenUrl) Then o.TokenUrl = tokenUrl

            o.TimeoutMilliseconds = ParseInt(GetValue(root, "TimeoutMilliseconds"), 30000)
            o.ReturnProbability = ParseBool(GetValue(root, "ReturnProbability"), True)
            o.ReturnLocation = ParseBool(GetValue(root, "ReturnLocation"), False)
            o.VerifyParameter = ParseBool(GetValue(root, "VerifyParameter"), False)
            o.MaxPages = ParseInt(GetValue(root, "MaxPages"), 1)
            o.PreferLocalParse = ParseBool(GetValue(root, "PreferLocalParse"), True)
            o.AutoFallbackToOcr = ParseBool(GetValue(root, "AutoFallbackToOcr"), True)
            Return o
        Catch ex As Exception
            AppLogger.Error("读取 OCR 外部配置失败（忽略）。", ex)
            Return Nothing
        End Try
    End Function

    Private Function GetValue(root As XElement, name As String) As String
        Dim el As XElement = root.Element(name)
        If el Is Nothing Then Return Nothing
        Return el.Value
    End Function

    Private Function ParseBool(v As String, fallback As Boolean) As Boolean
        Dim b As Boolean
        If Boolean.TryParse(v, b) Then Return b
        Return fallback
    End Function

    Private Function ParseInt(v As String, fallback As Integer) As Integer
        Dim i As Integer
        If Integer.TryParse(v, i) Then Return i
        Return fallback
    End Function

End Module
