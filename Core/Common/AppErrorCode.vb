''' <summary>
''' 统一错误代码。业务流程用它标识错误类型，再由 UserFriendlyMessageProvider 转为
''' 面向用户的简洁说明与建议操作。日志保留详细信息，弹窗只显示友好文案。
''' </summary>
Public Enum AppErrorCode
    None = 0

    ' —— 邮件/附件 ——
    NoMailSelected = 1          ' 未选择邮件
    SelectedNotMail = 2         ' 选中项不是邮件
    MailNoPdf = 3               ' 邮件无 PDF 附件
    AttachmentSaveFailed = 4    ' PDF 附件保存失败

    ' —— PDF ——
    PdfMergeFailed = 5          ' PDF 合并失败
    PdfTextExtractFailed = 6    ' PDF 文本抽取失败

    ' —— 识别 ——
    LocalFieldsInsufficient = 7 ' 本地识别字段不足
    OcrDisabled = 8             ' 百度 OCR 未启用
    OcrConfigMissing = 9        ' 百度 OCR 配置缺失
    OcrAuthFailed = 10          ' 百度 OCR 密钥错误
    OcrUnauthorizedOrQuota = 11 ' 未授权或额度不足
    OcrNetworkError = 12        ' 网络异常
    OcrTimeout = 13             ' 超时
    OcrEmptyFields = 14         ' 返回空字段
    OcrReturnedError = 15       ' 返回 error_code/error_msg

    ' —— 归档目录 ——
    ArchiveFolderNotConfigured = 16 ' 未配置归档目录
    ArchiveFolderMissing = 17       ' 归档目录不存在
    ArchiveFolderNotWritable = 18   ' 无写入权限

    ' —— 命名/文件 ——
    NamingTemplateEmpty = 19    ' 命名模板为空
    NamingUnknownVariable = 20  ' 命名模板含未知变量
    FileNameBuildFailed = 21    ' 文件名生成失败
    FileNameConflict = 22       ' 同名文件冲突
    FileCopyFailed = 23         ' 文件复制/移动失败

    ' —— 安全/配置/日志 ——
    DpapiDecryptFailed = 24     ' DPAPI 解密失败
    ConfigSaveFailed = 25       ' 配置保存失败
    LogWriteFailed = 26         ' 日志写入失败

    ' —— 其它 ——
    Unknown = 27                ' 未知异常

    ' —— 预检查补充 ——
    TempDirNotWritable = 30     ' 临时目录不可写
    LogDirNotWritable = 31      ' 日志目录不可写
    ArchiveAlreadyRunning = 32  ' 已有归档任务在运行
    ArchiveFolderPathInvalid = 33 ' 归档目录路径非法
    ArchivePathTooLong = 34     ' 归档完整路径可能超过 MAX_PATH
    ArchiveDiskSpaceLow = 35    ' 归档磁盘可用空间不足
End Enum
