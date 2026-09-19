Imports System.IO
Imports System.Globalization

''' <summary>
''' 百度智能财务票据识别识别器（真实调用）。
''' 流程：获取/缓存 Token → 逐页调用 multiple_invoice 接口 → 解析 JSON → 集中字段映射。
''' 约束：
'''  - 未启用 → ConfigurationMissing（"在线 OCR 未启用"）；配置缺失 → ConfigurationMissing；
'''  - 文件过大 → 明确失败；error_code → 失败并保留百度错误码/信息；
'''  - 成功但无字段 → 明确提示"识别成功但未提取到字段"；
'''  - 单页失败不影响其它页；token/AK/SK 不写日志。
''' </summary>
Public Class BaiduOcrInvoiceRecognizer
    Implements IInvoiceRecognizer

    ''' <summary>PDF 原始字节上限（超过则拒绝在线 OCR）。base64 后约 1.33 倍，留足余量。</summary>
    Private Const MaxRawPdfBytes As Integer = 6 * 1024 * 1024

    Private ReadOnly _options As BaiduOcrOptions
    Private ReadOnly _tokenProvider As BaiduAccessTokenProvider
    Private ReadOnly _httpClient As BaiduOcrHttpClient
    Private ReadOnly _parser As BaiduMultipleInvoiceResponseParser
    Private ReadOnly _mapper As BaiduInvoiceFieldMapper

    ''' <summary>必须显式传入配置（解耦 My.Settings，便于离线测试）。</summary>
    Public Sub New(options As BaiduOcrOptions)
        _options = If(options, New BaiduOcrOptions())
        _tokenProvider = New BaiduAccessTokenProvider()
        _httpClient = New BaiduOcrHttpClient()
        _parser = New BaiduMultipleInvoiceResponseParser()
        _mapper = New BaiduInvoiceFieldMapper()
    End Sub

    Public ReadOnly Property Name As String Implements IInvoiceRecognizer.Name
        Get
            Return "BaiduOcrInvoiceRecognizer"
        End Get
    End Property

    Public Function IsAvailable() As Boolean Implements IInvoiceRecognizer.IsAvailable
        ' 内网版本禁用在线解析
        If Not BuildFeatures.OnlineParserEnabled Then
            Return False
        End If
        Return _options IsNot Nothing AndAlso _options.IsConfigured()
    End Function

    Public Function Recognize(context As RecognitionContext) As InvoiceRecognitionResult Implements IInvoiceRecognizer.Recognize
        Dim result As New InvoiceRecognitionResult()
        result.RecognizerName = Name
        result.Source = RecognitionSource.BaiduOcr

        Try
            ' 内网版本禁用在线解析
            If Not BuildFeatures.OnlineParserEnabled Then
                result.Status = RecognitionStatus.ConfigurationMissing
                result.Message = "当前构建版本未启用在线解析功能。"
                Return result
            End If

            ' 未启用 / 配置缺失
            If _options Is Nothing OrElse Not _options.Enabled Then
                result.Status = RecognitionStatus.ConfigurationMissing
                result.Message = "在线 OCR 未启用。"
                Return result
            End If
            If Not _options.IsConfigured() Then
                result.Status = RecognitionStatus.ConfigurationMissing
                result.Message = "在线 OCR 配置缺失（请在设置中填写 API Key / Secret Key / 接口地址）。"
                Return result
            End If

            Dim pdfPath As String = If(context Is Nothing, Nothing, context.PdfPath)
            If String.IsNullOrEmpty(pdfPath) OrElse Not File.Exists(pdfPath) Then
                Return InvoiceRecognitionResult.Failure("待识别 PDF 不存在。")
            End If

            ' 先按文件长度判断，避免把超大 PDF 整个读入内存之后才拒绝（原实现顺序相反）。
            Dim pdfLength As Long = 0
            Try
                pdfLength = New FileInfo(pdfPath).Length
            Catch
            End Try
            If pdfLength > MaxRawPdfBytes Then
                Return InvoiceRecognitionResult.Failure(
                    String.Format(CultureInfo.InvariantCulture,
                                  "文件过大（{0:N0} 字节），无法在线 OCR（上限约 {1:N0} 字节）。", pdfLength, MaxRawPdfBytes))
            End If

            Dim pdfBytes As Byte() = File.ReadAllBytes(pdfPath)
            If pdfBytes.Length > MaxRawPdfBytes Then
                Return InvoiceRecognitionResult.Failure(
                    String.Format(CultureInfo.InvariantCulture,
                                  "文件过大（{0:N0} 字节），无法在线 OCR（上限约 {1:N0} 字节）。", pdfBytes.Length, MaxRawPdfBytes))
            End If

            ' 编码一次、跨页复用：多页识别时每页重算会对同一份 PDF 反复产生数十 MB 临时字符串。
            Dim encodedPdf As String = BaiduOcrHttpClient.EncodePdfForUpload(pdfBytes)

            ' 获取 Token（缓存）
            Dim tokenResult As BaiduAccessTokenResult = _tokenProvider.GetToken(_options)
            If Not tokenResult.Success Then
                Dim r As InvoiceRecognitionResult = InvoiceRecognitionResult.Failure(
                    "获取百度 Access Token 失败：" & tokenResult.ErrorMessage &
                    If(tokenResult.HttpStatusCode > 0, "（HTTP " & tokenResult.HttpStatusCode & "）", ""))
                r.Source = RecognitionSource.BaiduOcr
                Return r
            End If

            Dim maxPages As Integer = Math.Max(1, _options.MaxPages)
            Dim mappedAnyField As Boolean = False
            Dim pageErrors As New List(Of String)()
            Dim successfulPages As Integer = 0

            For page As Integer = 1 To maxPages
                Dim raw As BaiduOcrRawResponse = _httpClient.RecognizeMultipleInvoice(_options, tokenResult.AccessToken, encodedPdf, page)

                If Not raw.Success Then
                    pageErrors.Add("第 " & page & " 页请求失败：" & If(raw.NetworkError, "网络错误") & "（HTTP " & raw.HttpStatusCode & "）")
                    Continue For
                End If

                If raw.HasBaiduError Then
                    ' error_code/error_msg：记录并继续（单页失败不中断）
                    pageErrors.Add("第 " & page & " 页百度错误：error_code=" & raw.ErrorCode & ", error_msg=" & If(raw.ErrorMsg, ""))
                    AppLogger.Warn("百度 OCR 返回错误码：" & raw.ErrorCode & " / " & If(raw.ErrorMsg, ""))
                    Continue For
                End If

                successfulPages += 1
                If Not String.IsNullOrEmpty(raw.RawJson) AndAlso String.IsNullOrEmpty(result.RawText) Then
                    result.RawText = raw.RawJson
                End If

                Dim parsed As BaiduParsedDocument = _parser.Parse(raw.RawJson)
                If Not String.IsNullOrEmpty(parsed.PdfFileSize) Then
                    result.Invoice.SetExtendedField("pdf_file_size", parsed.PdfFileSize)
                End If

                If _mapper.Map(parsed, result) Then
                    mappedAnyField = True
                End If
            Next

            ' 汇总页级错误到消息
            For Each pe As String In pageErrors
                result.Messages.Add(pe)
            Next

            ' 结论
            If successfulPages = 0 Then
                Dim fail As InvoiceRecognitionResult = InvoiceRecognitionResult.Failure(
                    "百度 OCR 全部页失败。" & String.Join("；", pageErrors.ToArray()))
                fail.Source = RecognitionSource.BaiduOcr
                Return fail
            End If

            If Not mappedAnyField OrElse Not result.Invoice.HasAnyCoreField() Then
                result.Status = RecognitionStatus.PartialSuccess
                result.Message = "OCR 识别成功但未提取到字段。"
            Else
                result.MissingKeyFields = KeyFieldEvaluator.GetMissingKeyFields(result.Invoice, result.DocumentType)
                If result.MissingKeyFields.Count = 0 Then
                    result.Status = RecognitionStatus.Success
                    result.Message = "百度 OCR 识别成功。"
                Else
                    result.Status = RecognitionStatus.PartialSuccess
                    result.Message = "百度 OCR 部分成功，缺失：" & String.Join("、", result.MissingKeyFields.ToArray())
                End If
            End If

            AppLogger.Info(String.Format("百度 OCR 结果：状态={0}, 类型={1}, 字段数={2}, 成功页={3}",
                                         result.Status, result.DocumentType, result.Fields.Count, successfulPages))
            Return result

        Catch ex As Exception
            AppLogger.Error("百度 OCR 识别异常。", ex)
            Dim r As InvoiceRecognitionResult = InvoiceRecognitionResult.Failure("百度 OCR 识别异常：" & ExceptionFormatter.ToUserMessage(ex))
            r.Source = RecognitionSource.BaiduOcr
            Return r
        End Try
    End Function

End Class
