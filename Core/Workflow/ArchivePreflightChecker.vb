Imports System.IO
Imports System.Globalization

''' <summary>
''' 归档前预检查：在真正开始批量处理前，一次性检查配置与运行环境，避免中途才失败。
''' 设计为可离线测试（入参为配置值，不依赖 Outlook）；邮件选择数由调用方（Ribbon）传入。
'''
''' 注意：本类【不再】管理“归档任务是否正在运行”。运行锁由 <see cref="ArchiveRunGuard"/> 统一负责，
''' 预检查只检查配置、目录、选中邮件、权限等静态条件，避免与外层运行锁产生“自我阻断”。
''' （历史缺陷：外层先 TryBeginRun 置位，预检查又检查 IsRunning，导致每次点击都误报“已有任务正在运行”。）
''' </summary>
Public Class ArchivePreflightChecker

    ''' <summary>
    ''' 执行预检查。
    ''' </summary>
    ''' <param name="selectedCount">Outlook 当前选中项数量（由 Ribbon 传入；-1 表示未知/不检查）。</param>
    ''' <param name="archiveFolder">归档目录。</param>
    ''' <param name="ocrOptions">OCR 配置。</param>
    ''' <param name="unifiedTemplate">统一命名模板（可空 → 用默认）。</param>
    Public Function Check(selectedCount As Integer, archiveFolder As String, ocrOptions As BaiduOcrOptions, unifiedTemplate As String) As ArchivePreflightResult
        Dim r As New ArchivePreflightResult()

        ' 说明：明文 Secret Key 的 DPAPI 迁移（依赖 My.Settings）已移至 MainRibbon 在预检查之前执行，
        ' 使本预检查保持“纯配置值入参、不依赖 Outlook / My.Settings”，从而可离线单元测试。

        ' 说明：运行状态检查已移除。是否“已有归档任务正在运行”由 ArchiveRunGuard 在预检查之前判断，
        ' 预检查不再检查运行状态，避免与外层运行锁自我阻断。

        ' 选中邮件
        If selectedCount = 0 Then
            r.AddCode(AppErrorCode.NoMailSelected, blocking:=True)
        End If

        ' 归档目录
        CheckArchiveFolder(archiveFolder, r)

        ' 临时目录 / 日志目录
        Dim tempDir As String = Nothing
        Try
            tempDir = PathHelper.GetTempWorkDirectory()
        Catch
        End Try
        If String.IsNullOrEmpty(tempDir) OrElse Not IsDirWritable(tempDir) Then
            r.AddCode(AppErrorCode.TempDirNotWritable, blocking:=True)
        End If

        Dim logDir As String = Nothing
        Try
            logDir = PathHelper.GetLogDirectory(archiveFolder)
        Catch
        End Try
        If String.IsNullOrEmpty(logDir) OrElse Not IsDirWritable(logDir) Then
            r.AddCode(AppErrorCode.LogDirNotWritable, blocking:=False) ' 日志问题不阻断
        End If

        ' 命名模板
        If String.IsNullOrWhiteSpace(unifiedTemplate) Then
            r.AddCode(AppErrorCode.NamingTemplateEmpty, blocking:=False)
        End If

        ' OCR 配置（仅提示，不阻断——本地可能已足够）
        If ocrOptions IsNot Nothing AndAlso ocrOptions.Enabled AndAlso Not ocrOptions.IsConfigured() Then
            r.AddCode(AppErrorCode.OcrConfigMissing, blocking:=False)
        End If

        ' 汇总 + 逐条明细日志（避免以后只看到“问题数=1”而无法定位是哪一项）。
        AppLogger.Info("归档预检查完成：问题数=" & r.Issues.Count & "，阻断=" & r.HasBlocking)
        For Each issue As ArchivePreflightIssue In r.Issues
            AppLogger.Info(String.Format("  预检查{0}：Code={1}, Severity={2}, Blocking={3}, Message={4}",
                                         If(issue.IsBlocking, "[阻断]", "[提示]"),
                                         issue.Code, issue.Severity, issue.IsBlocking, issue.UserMessage))
        Next
        Return r
    End Function

    Private Sub CheckArchiveFolder(archiveFolder As String, r As ArchivePreflightResult)
        If String.IsNullOrWhiteSpace(archiveFolder) Then
            r.AddCode(AppErrorCode.ArchiveFolderNotConfigured, blocking:=True)
            Return
        End If

        ' 路径格式校验
        Dim full As String = Nothing
        Try
            full = Path.GetFullPath(archiveFolder)
        Catch
            r.AddCode(AppErrorCode.ArchiveFolderPathInvalid, blocking:=True)
            Return
        End Try

        ' 路径长度校验（O-12）：项目无 app.config，.NET 4.8 下未启用长路径，
        ' 超长路径会在 File.Copy 时抛 PathTooLongException —— 且只在批次跑到一半时才暴露。
        ' 这里按"最长可能产出名"做非阻断提示（并非每个文件都会达到最长）。
        If Not String.IsNullOrEmpty(full) Then
            Dim longestName As Integer = MaxBaseNameLength + ConflictSuffixReserve + ExtensionReserve
            If full.Length + 1 + longestName > MaxPathLength Then
                AppLogger.Warn("归档目录路径较长（" & full.Length.ToString(CultureInfo.InvariantCulture) &
                               " 字符），加上文件名后可能超过 " & MaxPathLength.ToString(CultureInfo.InvariantCulture) & "。")
                r.AddCode(AppErrorCode.ArchivePathTooLong, blocking:=False)
            End If
        End If

        ' 不存在 → 尝试创建（相当于“自动创建”）
        If Not SafeDirExists(archiveFolder) Then
            If Not PathHelper.EnsureDirectory(archiveFolder) Then
                r.AddCode(AppErrorCode.ArchiveFolderMissing, blocking:=True)
                Return
            End If
        End If

        ' 写权限
        If Not IsDirWritable(archiveFolder) Then
            r.AddCode(AppErrorCode.ArchiveFolderNotWritable, blocking:=True)
        End If

        ' 磁盘可用空间（O-16）：仅提示，不阻断。
        CheckFreeSpace(archiveFolder, r)
    End Sub

    ''' <summary>Windows 传统 MAX_PATH 上限（含结尾 NUL 共 260，可用路径长度 259）。</summary>
    Private Const MaxPathLength As Integer = 259

    ''' <summary>文件名主体最大长度（与 <see cref="FileNameSanitizer"/> 的 MaxBaseNameLength 保持一致）。</summary>
    Private Const MaxBaseNameLength As Integer = 120

    ''' <summary>"(nn)" 同名冲突后缀预留。</summary>
    Private Const ConflictSuffixReserve As Integer = 6

    ''' <summary>扩展名 ".pdf" 预留。</summary>
    Private Const ExtensionReserve As Integer = 4

    ''' <summary>归档磁盘可用空间下限（100 MB）：低于该值仅提示，不阻断。</summary>
    Private Const MinFreeSpaceBytes As Long = 100L * 1024 * 1024

    ''' <summary>
    ''' 磁盘可用空间提示（O-16）。单批需要为每张 PDF 复制一份，滴滴发票还需额外的合并副本，
    ''' 空间不足时会在批次中途逐项失败，因此提前给出提示。
    ''' 无法获取磁盘信息（UNC / 异常）时不作判定。
    ''' </summary>
    Private Sub CheckFreeSpace(folder As String, r As ArchivePreflightResult)
        Try
            Dim root As String = Path.GetPathRoot(Path.GetFullPath(folder))
            If String.IsNullOrWhiteSpace(root) Then Return
            Dim d As New DriveInfo(root)
            If Not d.IsReady Then Return
            If d.AvailableFreeSpace < MinFreeSpaceBytes Then
                AppLogger.Warn("归档磁盘可用空间不足：" &
                               (d.AvailableFreeSpace \ (1024 * 1024)).ToString(CultureInfo.InvariantCulture) & " MB（下限 " &
                               (MinFreeSpaceBytes \ (1024 * 1024)).ToString(CultureInfo.InvariantCulture) & " MB）")
                r.AddCode(AppErrorCode.ArchiveDiskSpaceLow, blocking:=False)
            End If
        Catch
            ' 无法获取时忽略。
        End Try
    End Sub

    Private Function SafeDirExists(dir As String) As Boolean
        Try
            Return Directory.Exists(dir)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 写权限检测：在目标目录创建并立即关闭一个探针文件。
    ''' 使用 <see cref="FileOptions.DeleteOnClose"/>：即使进程在写入后被强杀，内核在关闭句柄时
    ''' 也会删除该文件，因此不会在用户的归档目录里留下 .iwh_write_test_* 残留（O-16）。
    ''' </summary>
    Private Function IsDirWritable(dir As String) As Boolean
        Try
            If Not Directory.Exists(dir) Then Return False
            Dim probe As String = Path.Combine(dir, ".iwh_write_test_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ".tmp")
            Using fs As New FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.DeleteOnClose)
                fs.WriteByte(0)
            End Using
            Return True
        Catch
            Return False
        End Try
    End Function

End Class
