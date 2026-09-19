Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms
Imports System.Globalization

Public Class MainRibbon

    ''' <summary>
    ''' “归档”按钮：对当前选中邮件批量执行 读取PDF附件 → 识别 → 命名 → 归档，并展示简明汇总。
    ''' 单封失败不影响整批；详细过程写入日志；不逐个附件弹窗。
    '''
    ''' 启动顺序（防重复点击 + 防自我阻断）：
    '''   1) 先原子获取 ArchiveRunGuard 运行锁；获取失败 → 提示“已有任务正在运行”并返回；
    '''   2) 用 Using 包裹 token，保证任何路径（预检查失败/业务异常/进度窗口异常）都释放运行锁；
    '''   3) 执行归档前预检查（不再检查运行状态）；
    '''   4) 预检查失败 → 显示真实原因并返回（不创建进度窗口）；
    '''   5) 预检查通过 → 创建进度窗口 → 执行批量归档 → 显示汇总。
    ''' </summary>
    Private Sub ButtonArchive_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonArchive.Click
        Dim threadId As Integer = System.Threading.Thread.CurrentThread.ManagedThreadId
        AppLogger.Initialize(SafeSetting(Function() My.Settings.ArchiveFolderPath))
        AppLogger.Info("用户点击『归档』按钮。线程=" & threadId)

        Dim progress As ProgressForm = Nothing
        Try
            ' —— 第一步：原子获取唯一运行锁（防重复点击 / 防并发归档）——
            ' 直接用 Using 包裹"获取"动作本身，确保从获取成功那一刻起就没有
            ' "已持有但尚未进入 Using"的窗口（中间任何异常都会永久卡住进程级锁）。
            Using runToken As ArchiveRunToken = ArchiveRunGuard.TryAcquire("archive-" & DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture))
                If runToken Is Nothing Then
                    AppLogger.Warn("获取归档运行锁失败：已有归档任务正在运行。持有者：" & ArchiveRunGuard.DescribeHolder())
                    MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.ArchiveAlreadyRunning).ToUserText(),
                                    "工作助手 - 归档", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                AppLogger.Info("已获取归档运行锁。")

                ' 批次进行期间禁用归档按钮，避免 DoEvents 泵消息时用户重复触发
                ' （运行锁本身也会拒绝，但禁用按钮能给出更清晰的交互反馈）。
                Try
                    ButtonArchive.Enabled = False
                Catch
                End Try

                Dim application = Globals.ThisAddIn.Application

                ' 归档前迁移：确保明文 Secret Key 已转换为 DPAPI 加密（防止用户不打开设置直接点归档）。
                ' 依赖 My.Settings，故放在 Ribbon 侧执行；失败不阻断，仅记日志。
                Try
                    If ProtectedSettingsProvider.MigratePlaintextIfNeeded() Then
                        AppLogger.Info("归档前检测到明文 Secret Key 已迁移为 DPAPI 加密。")
                    End If
                Catch migEx As Exception
                    AppLogger.Warn("归档前密钥迁移失败（不阻断）：" & migEx.Message)
                End Try

                ' —— 归档前预检查（只检查配置/目录/选中邮件/权限；运行状态由运行锁负责）——
                Dim selectedCount As Integer = GetSelectedCount(application)
                Dim archiveFolder As String = SafeSetting(Function() My.Settings.ArchiveFolderPath)
                Dim ocrOptions As BaiduOcrOptions = OcrConfigProvider.LoadBaiduOptions()
                Dim template As String = SafeSetting(Function() My.Settings.UnifiedNameTemplate)

                AppLogger.Info("归档前预检查开始。选中邮件数=" & selectedCount)
                Dim checker As New ArchivePreflightChecker()
                Dim pre As ArchivePreflightResult = checker.Check(selectedCount, archiveFolder, ocrOptions, template)

                If pre.HasBlocking Then
                    AppLogger.Warn("归档前预检查未通过（存在阻断项），已取消本次归档，运行锁将随即释放。")
                    MessageBox.Show(pre.BuildUserText() & Environment.NewLine & "日志: " & AppLogger.CurrentLogPath(),
                                    "工作助手 - 无法归档", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return ' Using 释放运行锁；此路径未创建进度窗口
                End If
                AppLogger.Info("归档前预检查通过，开始批量归档。")

                ' 预检查通过后才创建进度窗口（无取消按钮，禁止关闭）。
                progress = New ProgressForm()
                progress.Show()
                System.Windows.Forms.Application.DoEvents()
                Dim reporter As New UiArchiveProgressReporter(progress)

                Dim workflow As New BatchArchiveWorkflow()
                Dim batch As ArchiveBatchResult = workflow.Run(application, reporter)
                AppLogger.Info("批量归档流程结束。批次=" & batch.BatchId)

                ' 批次因用户取消而提前结束时，在进度窗口上给出明确收尾提示。
                If reporter.IsCancellationRequested Then
                    Try
                        progress.MarkCanceled()
                    Catch
                    End Try
                End If

                If progress IsNot Nothing Then
                    Try
                        progress.Close()
                    Catch closeEx As Exception
                        ' 进度窗口收尾失败不应把已成功的归档报成"未知错误"。
                        AppLogger.Warn("关闭进度窗口失败（忽略）：" & closeEx.Message)
                    End Try
                    progress = Nothing
                End If

                ' —— 归档正常结束、且本批确有文件成功归档 → 打开/激活归档目录（整批仅一次）——
                ' 无成功归档文件则不打开；打开/激活失败仅记日志，不影响归档主流程完成状态。
                Try
                    If batch IsNot Nothing AndAlso batch.ArchivedFileCount > 0 Then
                        ExplorerFolderService.OpenOrActivateFolder(archiveFolder)
                    Else
                        AppLogger.Info("本批无成功归档文件，跳过打开归档目录。")
                    End If
                Catch openEx As Exception
                    AppLogger.Warn("打开归档目录调用异常（已忽略，不影响归档）：" & openEx.Message)
                End Try

                ShowSummary(batch)
            End Using ' 运行锁在此释放
        Catch ex As Exception
            AppLogger.Error("归档流程发生未处理异常。", ex)
            MessageBox.Show(UserFriendlyMessageProvider.Describe(AppErrorCode.Unknown).ToUserText() & Environment.NewLine &
                            "日志: " & AppLogger.CurrentLogPath(),
                            "工作助手 - 错误", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        Finally
            Try
                If progress IsNot Nothing AndAlso Not progress.IsDisposed Then
                    progress.Close()
                    progress.Dispose()
                End If
            Catch
            End Try
            ' 无论成功失败都恢复按钮可用状态。
            Try
                ButtonArchive.Enabled = True
            Catch
            End Try
            AppLogger.Info("归档运行锁已释放（线程=" & threadId & "）。")
        End Try
    End Sub

    ''' <summary>展示归档完成汇总（以邮件为单位，含报告与日志位置）。</summary>
    Private Sub ShowSummary(batch As ArchiveBatchResult)
        Select Case batch.OverallStatus
            Case ProcessStatus.ConfigurationMissing
                MessageBox.Show(FirstMessage(batch, UserFriendlyMessageProvider.Describe(AppErrorCode.ArchiveFolderNotConfigured).ToUserText()),
                                "无法归档", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Case ProcessStatus.Skipped
                MessageBox.Show(FirstMessage(batch, "没有可处理的邮件（无 PDF 附件）。请先选中含 PDF 附件的邮件。"),
                                "工作助手 - 归档", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Case Else
                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine(batch.BuildSummaryText())
                sb.AppendLine("需要配置 OCR 的邮件: " & batch.NeedsOcrCount)

                ' 面向用户的回退提示（简明、不含候选评分等技术细节）
                If CountOcrFallback(batch) > 0 Then
                    sb.AppendLine("部分发票本地识别字段不足，已自动使用在线 OCR 兜底。")
                End If
                If batch.NeedsOcrCount > 0 Then
                    If OcrConfiguredSafe() Then
                        sb.AppendLine("仍有部分发票字段不足且在线 OCR 未能补全，请人工核对命名。")
                    Else
                        sb.AppendLine("部分发票字段不足且未启用在线 OCR，请在设置中启用在线 OCR 以提升识别率。")
                    End If
                End If

                sb.AppendLine()
                If Not String.IsNullOrEmpty(batch.ReportPath) Then sb.AppendLine("结果报告: " & batch.ReportPath)
                sb.AppendLine("日志: " & AppLogger.CurrentLogPath())
                Dim icon As MessageBoxIcon = If(batch.FailureCount > 0, MessageBoxIcon.Warning, MessageBoxIcon.Information)
                MessageBox.Show(sb.ToString(), "工作助手 - 归档完成（批次 " & batch.BatchId & "）", MessageBoxButtons.OK, icon)
        End Select
    End Sub

    ''' <summary>统计本批中通过在线 OCR 兜底（来源 BaiduOcr / Mixed）的归档项数。</summary>
    Private Function CountOcrFallback(batch As ArchiveBatchResult) As Integer
        Dim n As Integer = 0
        If batch Is Nothing OrElse batch.Items Is Nothing Then Return 0
        For Each it As ArchiveItemResult In batch.Items
            If it.RecognitionSource = RecognitionSource.BaiduOcr OrElse it.RecognitionSource = RecognitionSource.Mixed Then n += 1
        Next
        Return n
    End Function

    ''' <summary>安全判断在线 OCR 是否已配置（失败按未配置处理）。</summary>
    Private Function OcrConfiguredSafe() As Boolean
        Try
            Dim o As BaiduOcrOptions = OcrConfigProvider.LoadBaiduOptions()
            Return o IsNot Nothing AndAlso o.IsConfigured()
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 安全获取当前选中项数量（COM，失败返回 -1）。
    ''' 必须显式释放 ActiveExplorer / Selection：Selection 会传递引用全部选中 MailItem，
    ''' 不释放会让邮件对象以非确定方式驻留（违反 MailAttachmentReader 自述的 COM 释放契约）。
    ''' </summary>
    Private Function GetSelectedCount(application As Object) As Integer
        Dim explorer As Object = Nothing
        Dim selection As Object = Nothing
        Try
            explorer = application.ActiveExplorer()
            If explorer Is Nothing Then Return 0
            selection = explorer.Selection
            If selection Is Nothing Then Return 0
            Return CInt(selection.Count)
        Catch
            Return -1 ' 未知：不因此阻断
        Finally
            ComRelease.ReleaseAll(selection, explorer)
        End Try
    End Function

    Private Function SafeSetting(getter As Func(Of String)) As String
        Try
            Return getter()
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 取批次首条消息作为提示文案。消息可能来自 MailAttachmentReader 并**包含原始邮件主题**，
    ''' 因此进入 UI 前必须脱敏，与日志路径的脱敏口径保持一致。
    ''' </summary>
    Private Function FirstMessage(batch As ArchiveBatchResult, fallback As String) As String
        If batch IsNot Nothing AndAlso batch.Messages IsNot Nothing AndAlso batch.Messages.Count > 0 Then
            Return PrivacySafeFormatter.MaskSubject(batch.Messages(0))
        End If
        Return fallback
    End Function

    Private Sub ButtonSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSettings.Click
        Using form As New SettingsForm()
            form.ShowDialog()
        End Using
    End Sub

End Class
