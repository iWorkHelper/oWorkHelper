''' <summary>
''' 错误码 → 面向用户的友好文案与建议操作的单一映射来源。
''' 所有弹窗/汇总/预检查提示均由此产生，保证文案一致、可理解、可操作，且不泄漏敏感信息。
''' </summary>
Public Module UserFriendlyMessageProvider

    Private Function Make(sev As ErrorSeverity, msg As String, suggestion As String) As AppError
        Return New AppError With {.Severity = sev, .UserMessage = msg, .Suggestion = suggestion}
    End Function

    ''' <summary>返回该错误码的级别 / 用户说明 / 建议（不含 DetailForLog）。</summary>
    Public Function Describe(code As AppErrorCode) As AppError
        Dim e As AppError
        Select Case code
            Case AppErrorCode.NoMailSelected
                e = Make(ErrorSeverity.Warning, "未选择邮件。", "请先在 Outlook 中选择需要归档的邮件。")
            Case AppErrorCode.SelectedNotMail
                e = Make(ErrorSeverity.Info, "选中项不是邮件，已跳过。", "请选择邮件项。")
            Case AppErrorCode.MailNoPdf
                e = Make(ErrorSeverity.Info, "当前邮件没有 PDF 附件，已跳过。", "确认邮件中包含 PDF 附件。")
            Case AppErrorCode.AttachmentSaveFailed
                e = Make(ErrorSeverity.Error, "PDF 附件保存失败。", "请检查磁盘空间与临时目录权限后重试。")
            Case AppErrorCode.PdfMergeFailed
                e = Make(ErrorSeverity.Error, "PDF 合并失败，已跳过该邮件，其他邮件继续处理。", "确认 PDF 未加密/未损坏。")
            Case AppErrorCode.PdfTextExtractFailed
                e = Make(ErrorSeverity.Warning, "PDF 文本抽取失败。", "该 PDF 可能为图片型，可启用在线 OCR 兜底。")
            Case AppErrorCode.LocalFieldsInsufficient
                e = Make(ErrorSeverity.Warning, "本地识别未取到足够的命名字段。", "建议在设置中启用在线 OCR 兜底。")
            Case AppErrorCode.OcrDisabled
                e = Make(ErrorSeverity.Warning, "百度 OCR 未启用，本地识别字段不足，无法完成归档。", "在设置中勾选启用在线 OCR 并填写 API Key / Secret Key。")
            Case AppErrorCode.OcrConfigMissing
                e = Make(ErrorSeverity.Error, "百度 OCR 配置不完整。", "请在设置中填写 API Key 和 Secret Key，并确认接口地址。")
            Case AppErrorCode.OcrAuthFailed
                e = Make(ErrorSeverity.Error, "百度 OCR 密钥无效。", "请核对 API Key / Secret Key 是否正确；如已重置请更新设置。")
            Case AppErrorCode.OcrUnauthorizedOrQuota
                e = Make(ErrorSeverity.Error, "百度 OCR 未授权或额度不足。", "请在百度智能云控制台确认接口已开通且额度充足。")
            Case AppErrorCode.OcrNetworkError
                e = Make(ErrorSeverity.Error, "百度 OCR 网络异常。", "请检查网络连接或代理设置后重试。")
            Case AppErrorCode.OcrTimeout
                e = Make(ErrorSeverity.Error, "百度 OCR 请求超时。", "可在设置中适当增大超时时间，或稍后重试。")
            Case AppErrorCode.OcrEmptyFields
                e = Make(ErrorSeverity.Warning, "在线 OCR 识别成功但未提取到字段。", "该票据版式可能不受支持；文件仍会归档，可人工核对。")
            Case AppErrorCode.OcrReturnedError
                e = Make(ErrorSeverity.Error, "百度 OCR 返回错误。", "请稍后重试；如反复出现请核对接口与账号。")
            Case AppErrorCode.ArchiveFolderNotConfigured
                e = Make(ErrorSeverity.Critical, "尚未配置归档目录。", "请先在设置中选择归档文件夹。")
            Case AppErrorCode.ArchiveFolderMissing
                e = Make(ErrorSeverity.Critical, "归档目录不存在。", "请在设置中选择有效目录，或允许自动创建。")
            Case AppErrorCode.ArchiveFolderNotWritable
                e = Make(ErrorSeverity.Critical, "归档目录无写入权限。", "请更换目录或检查该目录的写入权限。")
            Case AppErrorCode.ArchiveFolderPathInvalid
                e = Make(ErrorSeverity.Critical, "归档目录路径格式非法。", "请重新选择一个有效的本地目录。")
            Case AppErrorCode.ArchivePathTooLong
                e = Make(ErrorSeverity.Warning, "归档目录过深，加上文件名后可能超出 Windows 路径长度上限（260）。", "建议改用更短的归档目录（例如放在盘符根目录下的一级文件夹），或缩短命名模板。")
            Case AppErrorCode.ArchiveDiskSpaceLow
                e = Make(ErrorSeverity.Warning, "归档目录所在磁盘可用空间不足。", "请清理磁盘空间后重试；空间不足会导致部分文件归档失败。")
            Case AppErrorCode.NamingTemplateEmpty
                e = Make(ErrorSeverity.Info, "命名模板为空，将使用默认模板。", "如需自定义，请在设置中填写命名规则。")
            Case AppErrorCode.NamingUnknownVariable
                e = Make(ErrorSeverity.Warning, "命名模板中包含未知变量，这些变量不会被替换。", "请点变量说明核对可用变量。")
            Case AppErrorCode.FileNameBuildFailed
                e = Make(ErrorSeverity.Error, "生成归档文件名失败。", "已使用回退命名；可检查命名模板是否异常。")
            Case AppErrorCode.FileNameConflict
                e = Make(ErrorSeverity.Info, "存在同名文件，已自动追加序号。", "")
            Case AppErrorCode.FileCopyFailed
                e = Make(ErrorSeverity.Error, "归档文件复制失败。", "请检查目标目录权限与磁盘空间。")
            Case AppErrorCode.DpapiDecryptFailed
                e = Make(ErrorSeverity.Warning, "已保存的密钥无法在当前 Windows 用户下解密。", "请在设置中重新输入 Secret Key 并保存。")
            Case AppErrorCode.ConfigSaveFailed
                e = Make(ErrorSeverity.Error, "配置保存失败。", "请重试；如仍失败请检查用户配置目录权限。")
            Case AppErrorCode.LogWriteFailed
                e = Make(ErrorSeverity.Warning, "日志写入失败（不影响归档）。", "可检查日志目录权限。")
            Case AppErrorCode.TempDirNotWritable
                e = Make(ErrorSeverity.Critical, "临时工作目录不可写。", "请检查 %AppData%\iWorkHelper\temp 目录权限。")
            Case AppErrorCode.LogDirNotWritable
                e = Make(ErrorSeverity.Warning, "日志目录不可写（不影响归档）。", "可检查日志目录权限。")
            Case AppErrorCode.ArchiveAlreadyRunning
                e = Make(ErrorSeverity.Warning, "已有归档任务正在运行。", "请等待当前任务完成后再试。")
            Case Else
                e = Make(ErrorSeverity.Error, "发生未知错误。", "请重试；如反复出现请查看日志或联系维护者。")
        End Select
        e.Code = code
        Return e
    End Function

    ''' <summary>
    ''' 把百度 OCR 的失败消息/错误码归类为具体 AppErrorCode（不含敏感信息）。
    ''' 优先依据百度 error_code（最精确），其次依据 HTTP 状态码，最后回退到消息关键词。
    ''' </summary>
    Public Function ClassifyOcrFailure(message As String, Optional httpStatus As Integer = 0, Optional baiduErrorCode As String = Nothing) As AppErrorCode
        Dim m As String = If(message, "").ToLowerInvariant()

        Dim code As Integer = 0
        If Not String.IsNullOrWhiteSpace(baiduErrorCode) Then Integer.TryParse(baiduErrorCode.Trim(), code)

        ' 百度错误码优先：110/111 = access_token 无效/过期；17/18/19 = 配额或 QPS 限制。
        If code = 110 OrElse code = 111 Then Return AppErrorCode.OcrAuthFailed
        If code = 17 OrElse code = 18 OrElse code = 19 Then Return AppErrorCode.OcrUnauthorizedOrQuota

        If httpStatus = 401 OrElse m.Contains("authentication") OrElse m.Contains("invalid_client") OrElse m.Contains("密钥") Then
            Return AppErrorCode.OcrAuthFailed
        End If
        If httpStatus = 403 OrElse m.Contains("unauthorized") OrElse m.Contains("quota") OrElse m.Contains("limit") OrElse m.Contains("额度") Then
            Return AppErrorCode.OcrUnauthorizedOrQuota
        End If
        If m.Contains("timeout") OrElse m.Contains("超时") OrElse m.Contains("timed out") Then
            Return AppErrorCode.OcrTimeout
        End If
        ' 注意：此处原先用 m.Contains("uri") 判定网络错误，但 "security" 也包含 "uri"，
        ' 会把 TLS/安全通道错误误报为网络错误。改为明确的关键词。
        If m.Contains("网络") OrElse m.Contains("network") OrElse m.Contains("远程") OrElse m.Contains("连接") OrElse
           m.Contains("安全通道") OrElse m.Contains("tls") OrElse m.Contains("ssl") OrElse m.Contains("certificate") Then
            Return AppErrorCode.OcrNetworkError
        End If
        Return AppErrorCode.OcrReturnedError
    End Function

End Module
