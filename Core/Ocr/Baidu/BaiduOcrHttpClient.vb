Imports System.IO
Imports System.Globalization
Imports System.Net
Imports System.Text
Imports System.Web.Script.Serialization

''' <summary>
''' 百度 OCR HTTP 客户端。使用 HttpWebRequest（不引第三方库）调用智能财务票据识别接口。
''' 请求规范：
'''  - POST，Content-Type: application/x-www-form-urlencoded；
'''  - URL 参数 access_token；
'''  - Body 参数 pdf_file（PDF 二进制 Base64 后再 UrlEncode）、pdf_file_num（页码）；
'''  - 可选 probability / location / verify_parameter 按配置传入；
'''  - 不使用 image / url，不处理图片附件。
''' </summary>
Public Class BaiduOcrHttpClient

    ''' <summary>
    ''' 在 .NET Framework 4.8 下显式启用 TLS 1.2（叠加，不覆盖已有协议）。
    ''' </summary>
    Public Shared Sub EnsureTls12()
        Try
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12
        Catch
            ' 某些环境不支持设置时忽略。
        End Try
    End Sub

    ''' <summary>
    ''' 把 PDF 字节编码为可放入表单体的 pdf_file 值：Base64 后再 UrlEncode。
    ''' 由调用方在**页循环之外**调用一次并跨页复用，避免同一份 PDF 被反复编码。
    ''' 注意：Uri.EscapeDataString 在 .NET Framework 下对超长字符串（约 65520 字符）会抛
    ''' "URI 字符串太长"；PDF 的 Base64 远超该上限，因此必须使用 WebUtility.UrlEncode。
    ''' </summary>
    Public Shared Function EncodePdfForUpload(pdfBytes As Byte()) As String
        If pdfBytes Is Nothing OrElse pdfBytes.Length = 0 Then Return String.Empty
        Return WebUtility.UrlEncode(Convert.ToBase64String(pdfBytes))
    End Function

    ''' <summary>
    ''' 调用智能财务票据识别接口，识别指定页码。
    ''' </summary>
    ''' <param name="options">OCR 配置。</param>
    ''' <param name="accessToken">已获取的 access_token（不写日志）。</param>
    ''' <param name="encodedPdfBase64">
    ''' 已 UrlEncode 的 PDF Base64（见 <see cref="EncodePdfForUpload"/>）。
    ''' 由调用方一次性计算后跨页复用：多页识别时若每页重算，
    ''' 会对同一份 PDF 反复产生数十 MB 的临时字符串。
    ''' </param>
    ''' <param name="pageNum">pdf_file_num 页码（从 1 开始）。</param>
    Public Function RecognizeMultipleInvoice(options As BaiduOcrOptions, accessToken As String,
                                            encodedPdfBase64 As String, pageNum As Integer) As BaiduOcrRawResponse
        Dim resp As New BaiduOcrRawResponse With {.PageIndex = pageNum}

        Try
            EnsureTls12()

            Dim url As String = options.OcrApiUrl & "?access_token=" & Uri.EscapeDataString(accessToken)

            Dim sb As New StringBuilder()
            sb.Append("pdf_file=").Append(If(encodedPdfBase64, String.Empty))
            sb.Append("&pdf_file_num=").Append(pageNum.ToString(CultureInfo.InvariantCulture))
            sb.Append("&probability=").Append(If(options.ReturnProbability, "true", "false"))
            sb.Append("&location=").Append(If(options.ReturnLocation, "true", "false"))
            sb.Append("&verify_parameter=").Append(If(options.VerifyParameter, "true", "false"))

            Dim bodyBytes As Byte() = Encoding.UTF8.GetBytes(sb.ToString())

            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            request.Timeout = Math.Max(5000, options.TimeoutMilliseconds)
            request.ContentLength = bodyBytes.Length

            Using reqStream As Stream = request.GetRequestStream()
                reqStream.Write(bodyBytes, 0, bodyBytes.Length)
            End Using

            Try
                Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                    resp.HttpStatusCode = CInt(response.StatusCode)
                    resp.RawJson = ReadStream(response.GetResponseStream())
                End Using
            Catch wex As WebException
                resp.HttpStatusCode = ExtractStatus(wex)
                resp.RawJson = ReadWebExceptionBody(wex)
                If String.IsNullOrEmpty(resp.RawJson) Then
                    resp.Success = False
                    resp.NetworkError = wex.Message
                    AppLogger.Warn("OCR 请求网络错误（页 " & pageNum & "）：" & wex.Message & "，HTTP=" & resp.HttpStatusCode)
                    Return resp
                End If
            End Try

            resp.Success = True
            ExtractBaiduError(resp)
            AppLogger.Info("OCR 请求完成（页 " & pageNum & "）：HTTP=" & resp.HttpStatusCode &
                           If(resp.HasBaiduError, "，error_code=" & resp.ErrorCode, ""))
            Return resp

        Catch ex As Exception
            resp.Success = False
            resp.NetworkError = ExceptionFormatter.ToUserMessage(ex)
            AppLogger.Error("OCR 请求异常（页 " & pageNum & "）。", ex)
            Return resp
        End Try
    End Function

    ''' <summary>从原始 JSON 中提取 error_code / error_msg / pdf_file_size（如有）。</summary>
    Private Sub ExtractBaiduError(resp As BaiduOcrRawResponse)
        Try
            If String.IsNullOrEmpty(resp.RawJson) Then Return
            Dim serializer As New JavaScriptSerializer()
            Dim map As Dictionary(Of String, Object) = TryCast(serializer.DeserializeObject(resp.RawJson), Dictionary(Of String, Object))
            If map Is Nothing Then Return
            If map.ContainsKey("error_code") AndAlso map("error_code") IsNot Nothing Then
                resp.ErrorCode = Convert.ToString(map("error_code"))
            End If
            If map.ContainsKey("error_msg") AndAlso map("error_msg") IsNot Nothing Then
                resp.ErrorMsg = Convert.ToString(map("error_msg"))
            End If
        Catch
            ' 提取失败不影响主流程，交由解析器进一步处理。
        End Try
    End Sub

    ''' <summary>响应体上限（16 MB）：防止异常/恶意响应造成无界内存增长。</summary>
    Private Const MaxResponseBytes As Integer = 16 * 1024 * 1024

    ''' <summary>
    ''' 读取响应体。带字节上限，超限抛 InvalidOperationException
    ''' （由上层转为可处理的失败，而不是把进程内存吃满）。
    ''' </summary>
    Private Function ReadStream(s As Stream) As String
        If s Is Nothing Then Return String.Empty
        Dim buffer(8191) As Byte
        Dim total As Integer = 0
        Using ms As New MemoryStream()
            Dim read As Integer = s.Read(buffer, 0, buffer.Length)
            While read > 0
                total += read
                If total > MaxResponseBytes Then
                    Throw New InvalidOperationException(
                        "OCR 响应体超过上限（" & MaxResponseBytes.ToString(CultureInfo.InvariantCulture) & " 字节），已中止读取。")
                End If
                ms.Write(buffer, 0, read)
                read = s.Read(buffer, 0, buffer.Length)
            End While
            Return Encoding.UTF8.GetString(ms.ToArray())
        End Using
    End Function

    Private Function ReadWebExceptionBody(wex As WebException) As String
        Try
            If wex.Response IsNot Nothing Then
                Return ReadStream(wex.Response.GetResponseStream())
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Function ExtractStatus(wex As WebException) As Integer
        Try
            Dim resp As HttpWebResponse = TryCast(wex.Response, HttpWebResponse)
            If resp IsNot Nothing Then
                Return CInt(resp.StatusCode)
            End If
        Catch
        End Try
        Return 0
    End Function

End Class
