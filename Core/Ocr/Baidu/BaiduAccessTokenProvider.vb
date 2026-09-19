Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web.Script.Serialization

''' <summary>
''' 百度 Access Token 提供者：使用 AK/SK 换取 access_token，并按 expires_in 缓存。
''' 特性：
'''  - 缓存 token，避免每个 PDF 重新获取；
'''  - 过期前预留安全缓冲（默认提前 5 分钟）刷新；
'''  - 显式启用 TLS 1.2（.NET Framework 4.8）；
'''  - 使用 JavaScriptSerializer 解析返回；
'''  - 失败返回结构化错误，不抛到 UI；token 绝不写日志。
''' 缓存为进程内共享（Shared），按 AK 区分。
''' </summary>
Public Class BaiduAccessTokenProvider

    Private Const RefreshBufferSeconds As Integer = 300 ' 提前 5 分钟刷新

    Private Shared ReadOnly SyncRoot As New Object()
    Private Shared _cachedToken As String
    Private Shared _cachedForApiKey As String
    Private Shared _expiresAtUtc As DateTime = DateTime.MinValue

    ''' <summary>
    ''' 获取有效 access_token（命中缓存则直接返回）。
    ''' </summary>
    Public Function GetToken(options As BaiduOcrOptions) As BaiduAccessTokenResult
        If options Is Nothing OrElse String.IsNullOrWhiteSpace(options.ApiKey) OrElse String.IsNullOrWhiteSpace(options.SecretKey) Then
            Return BaiduAccessTokenResult.Fail("缺少 API Key / Secret Key，无法获取 Access Token。")
        End If

        SyncLock SyncRoot
            ' 命中缓存（同一 AK 且未进入刷新缓冲期）
            If Not String.IsNullOrEmpty(_cachedToken) _
               AndAlso String.Equals(_cachedForApiKey, options.ApiKey, StringComparison.Ordinal) _
               AndAlso DateTime.UtcNow < _expiresAtUtc Then
                Return BaiduAccessTokenResult.Ok(_cachedToken, CInt((_expiresAtUtc - DateTime.UtcNow).TotalSeconds))
            End If
        End SyncLock

        ' 缓存未命中 → 请求新 token
        Dim fetched As BaiduAccessTokenResult = FetchToken(options)
        If fetched.Success Then
            SyncLock SyncRoot
                _cachedToken = fetched.AccessToken
                _cachedForApiKey = options.ApiKey
                Dim ttl As Integer = Math.Max(60, fetched.ExpiresInSeconds - RefreshBufferSeconds)
                _expiresAtUtc = DateTime.UtcNow.AddSeconds(ttl)
            End SyncLock
            AppLogger.Info("已获取百度 Access Token（有效期约 " & fetched.ExpiresInSeconds.ToString() & " 秒，已缓存）。")
        End If
        Return fetched
    End Function

    ''' <summary>清空缓存（如密钥变更/失效时可调用）。</summary>
    Public Shared Sub InvalidateCache()
        SyncLock SyncRoot
            _cachedToken = Nothing
            _cachedForApiKey = Nothing
            _expiresAtUtc = DateTime.MinValue
        End SyncLock
    End Sub

    Private Function FetchToken(options As BaiduOcrOptions) As BaiduAccessTokenResult
        Try
            BaiduOcrHttpClient.EnsureTls12()

            ' 凭据放在 POST body，而不是 URL 查询串：
            ' client_secret 出现在查询串中极易被代理日志、HTTP 访问日志与错误遥测记录。
            ' 百度 OAuth 接口同时支持查询串与表单体两种传参方式。
            Dim url As String = options.TokenUrl
            Dim form As String = "grant_type=client_credentials" &
                                 "&client_id=" & Uri.EscapeDataString(options.ApiKey) &
                                 "&client_secret=" & Uri.EscapeDataString(options.SecretKey)
            Dim formBytes As Byte() = Encoding.UTF8.GetBytes(form)

            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            request.Timeout = Math.Max(5000, options.TimeoutMilliseconds)
            request.ReadWriteTimeout = Math.Max(5000, options.TimeoutMilliseconds)
            request.ContentLength = formBytes.Length

            Using reqStream As System.IO.Stream = request.GetRequestStream()
                reqStream.Write(formBytes, 0, formBytes.Length)
            End Using

            Dim statusCode As Integer = 0
            Dim body As String = Nothing
            Try
                Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                    statusCode = CInt(response.StatusCode)
                    body = ReadResponse(response)
                End Using
            Catch wex As WebException
                statusCode = ExtractStatus(wex)
                body = ReadWebExceptionBody(wex)
                If String.IsNullOrEmpty(body) Then
                    Return BaiduAccessTokenResult.Fail("Token 请求网络错误：" & wex.Message, statusCode)
                End If
            End Try

            Return ParseTokenResponse(body, statusCode)

        Catch ex As Exception
            AppLogger.Error("获取百度 Access Token 异常。", ex)
            Return BaiduAccessTokenResult.Fail("获取 Access Token 异常：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

    Private Function ParseTokenResponse(body As String, statusCode As Integer) As BaiduAccessTokenResult
        Try
            Dim serializer As New JavaScriptSerializer()
            Dim map As Dictionary(Of String, Object) = TryCast(serializer.DeserializeObject(body), Dictionary(Of String, Object))
            If map Is Nothing Then
                Return BaiduAccessTokenResult.Fail("Token 返回无法解析。", statusCode)
            End If

            If map.ContainsKey("access_token") AndAlso map("access_token") IsNot Nothing Then
                Dim token As String = Convert.ToString(map("access_token"))
                Dim expiresIn As Integer = 2592000
                If map.ContainsKey("expires_in") AndAlso map("expires_in") IsNot Nothing Then
                    Integer.TryParse(Convert.ToString(map("expires_in")), expiresIn)
                End If
                Return BaiduAccessTokenResult.Ok(token, expiresIn)
            End If

            ' 错误：{"error":"invalid_client","error_description":"..."}
            Dim err As String = If(map.ContainsKey("error"), Convert.ToString(map("error")), Nothing)
            Dim errDesc As String = If(map.ContainsKey("error_description"), Convert.ToString(map("error_description")), Nothing)
            AppLogger.Warn("Token 获取失败：error=" & If(err, "(无)") & ", desc=" & If(errDesc, "(无)") & ", http=" & statusCode.ToString())
            Return BaiduAccessTokenResult.Fail("Token 获取失败：" & If(errDesc, If(err, "未知错误")), statusCode, err)

        Catch ex As Exception
            Return BaiduAccessTokenResult.Fail("解析 Token 返回异常：" & ExceptionFormatter.ToUserMessage(ex), statusCode)
        End Try
    End Function

    Private Function ReadResponse(response As HttpWebResponse) As String
        Using sr As New StreamReader(response.GetResponseStream(), Encoding.UTF8)
            Return sr.ReadToEnd()
        End Using
    End Function

    Private Function ReadWebExceptionBody(wex As WebException) As String
        Try
            If wex.Response IsNot Nothing Then
                Using sr As New StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8)
                    Return sr.ReadToEnd()
                End Using
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
