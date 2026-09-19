Imports System.IO
Imports System.Collections.Generic
Imports System.Globalization
Imports Outlook = Microsoft.Office.Interop.Outlook

''' <summary>
''' 批量归档工作流编排器（**按邮件/PDF 类型分流**）：
'''  - 滴滴发票邮件：发票 PDF + 行程单 PDF 合并为一个 PDF，统一命名归档；
'''  - 常规发票邮件：每个发票 PDF 单独识别、命名（常规发票模板）、归档；
'''  - 未识别 PDF：每个按 未识别_{原始文件名} 归档；
'''  - 无 PDF 邮件：跳过（不报错，计入跳过）。
''' 进度以邮件为主维度；单封邮件/单个 PDF 失败不影响其它；全程日志 + 报告。
''' </summary>
Public Class BatchArchiveWorkflow

    Private _carryNote As String = Nothing

    ''' <summary>本批次是否被用户取消（在邮件/PDF 边界置位）。</summary>
    Private _canceled As Boolean = False

    ' 共享服务/上下文（Run 内初始化，供分流方法复用）
    Private _archiveFolder As String
    Private _tempDir As String
    Private _runToken As String
    Private _generalTemplate As String
    Private _pipeline As RecognitionPipeline
    Private _merger As InvoiceRecognitionMerger
    Private _mergeSvc As PdfMergeService
    Private _executor As ArchiveExecutor
    Private _planner As ArchivePlanner
    Private _classifier As MailProcessingClassifier
    Private _ocrEnabled As Boolean
    Private _total As Integer

    ''' <summary>无进度上报的重载（离线/兼容）。</summary>
    Public Function Run(application As Outlook.Application) As ArchiveBatchResult
        Return Run(application, New NullArchiveProgressReporter())
    End Function

    ''' <summary>执行批量归档（分流），并上报进度。</summary>
    Public Function Run(application As Outlook.Application, reporter As IArchiveProgressReporter) As ArchiveBatchResult
        If reporter Is Nothing Then reporter = New NullArchiveProgressReporter()
        Dim batch As New ArchiveBatchResult()

        _archiveFolder = SafeSetting(Function() My.Settings.ArchiveFolderPath)
        Dim ocrOptions As BaiduOcrOptions = OcrConfigProvider.LoadBaiduOptions()
        AppLogger.Initialize(_archiveFolder)
        AppLogger.Info("=== 批量归档开始（分流）=== " & ocrOptions.ToSafeSummary())

        Dim templates As New NamingTemplates()
        templates.UnifiedTemplate = SafeSetting(Function() My.Settings.UnifiedNameTemplate)
        templates.GeneralInvoiceTemplate = SafeSetting(Function() My.Settings.GeneralInvoiceNameTemplate)
        _generalTemplate = If(String.IsNullOrWhiteSpace(templates.GeneralInvoiceTemplate), NamingTemplates.DefaultGeneralInvoiceTemplate, templates.GeneralInvoiceTemplate)

        _planner = New ArchivePlanner(templates)
        Dim validate As Result = _planner.ValidateArchiveFolder(_archiveFolder)
        If Not validate.IsSuccess Then
            batch.OverallStatus = validate.Status
            batch.AddMessage(validate.Message)
            AppLogger.Warn("归档目录校验失败：" & validate.Message)
            Return batch
        End If

        SafeReport(reporter, New ArchiveProgressInfo With {.Stage = ArchiveStage.Reading})
        Dim reader As New MailAttachmentReader()
        Dim grouping As MailPdfGroupingResult = reader.ReadSelectedPdfGroups(application)
        For Each m As String In grouping.Messages
            batch.AddMessage(m)
        Next
        If grouping.Groups.Count = 0 Then
            batch.OverallStatus = ProcessStatus.Skipped
            batch.AddMessage("没有可处理的邮件（无 PDF 附件）。")
            AppLogger.Info("无可处理邮件。")
            Return batch
        End If

        _pipeline = New RecognitionPipeline(ocrOptions)
        _merger = New InvoiceRecognitionMerger()
        _mergeSvc = New PdfMergeService()
        _executor = New ArchiveExecutor()
        _classifier = New MailProcessingClassifier()
        _tempDir = PathHelper.GetTempWorkDirectory()
        PurgeOldTempFiles()
        _runToken = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
        batch.BatchId = "B" & _runToken
        Dim reportStamp As String = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)
        _ocrEnabled = ocrOptions.IsConfigured()
        _total = grouping.Groups.Count
        AppLogger.Info("批次 ID=" & batch.BatchId & "，待处理邮件=" & _total)

        Dim processed As Integer = 0
        Try
            processed = ProcessGroups(batch, grouping, reporter)
        Finally
            ' 即使中途抛出异常，也必须产出报告与结束日志：
            ' 否则本批已累积的逐项结果全部丢失，用户只能看到一条通用错误。
            If _canceled Then
                batch.AddMessage("用户已取消：剩余邮件未处理，已归档的文件保留。")
                AppLogger.Warn("批次被用户取消，剩余邮件未处理。")
            End If
            batch.OverallStatus = ComputeOverall(batch)
            SafeReport(reporter, New ArchiveProgressInfo With {.TotalEmails = _total, .ProcessedEmails = processed, .Stage = ArchiveStage.Completed})
            batch.ReportPath = ArchiveReportWriter.Write(batch, _archiveFolder, reportStamp)
            AppLogger.Info("=== 批量归档结束（批次 " & batch.BatchId & "）：" & batch.BuildSummaryText() & " ===")
        End Try
        Return batch
    End Function

    ''' <summary>遍历选中邮件分组并逐封处理，返回已处理邮件数。单封失败不影响其它。</summary>
    Private Function ProcessGroups(batch As ArchiveBatchResult, grouping As MailPdfGroupingResult,
                                   reporter As IArchiveProgressReporter) As Integer
        Dim processed As Integer = 0
        For Each group As MailPdfGroup In grouping.Groups
            ' 在"邮件边界"响应取消：不在步骤中途强杀，避免留下半成品文件。
            If reporter IsNot Nothing AndAlso reporter.IsCancellationRequested Then
                _canceled = True
                Exit For
            End If
            Try
                ' 逐 PDF 识别 + 分类
                ReportStage(reporter, processed, group, ArchiveStage.ExtractingText)
                Dim perPdf As New List(Of PdfAttachmentClassificationResult)()
                For Each pdf As MailAttachmentItem In group.Pdfs
                    ' 识别（含 PDF 解析与可能的在线 OCR）是单批次里最耗时的步骤，
                    ' 因此在每个 PDF 之前也检查一次取消。
                    If reporter IsNot Nothing AndAlso reporter.IsCancellationRequested Then
                        _canceled = True
                        Exit For
                    End If
                    If _ocrEnabled Then ReportStage(reporter, processed, group, ArchiveStage.CallingOcr)
                    Dim recog As InvoiceRecognitionResult = _pipeline.Recognize(pdf.TempFilePath, pdf.OriginalFileName)
                    Dim cls As PdfAttachmentClassification = _classifier.ClassifyPdf(recog, recog.RawText, pdf.OriginalFileName, group.MailSubject)
                    perPdf.Add(New PdfAttachmentClassificationResult With {.Attachment = pdf, .Classification = cls, .Recognition = recog})
                Next

                If _canceled Then Exit For

                ' 邮件类型
                ReportStage(reporter, processed, group, ArchiveStage.Classifying)
                Dim classes As New List(Of PdfAttachmentClassification)()
                For Each pc As PdfAttachmentClassificationResult In perPdf
                    classes.Add(pc.Classification)
                Next
                Dim mailType As MailProcessingType = _classifier.ClassifyMail(classes)
                AppLogger.Info(String.Format("邮件[{0}]《{1}》类型={2}，PDF数={3}",
                                             group.Index, PrivacySafeFormatter.MaskSubject(group.MailSubject), mailType, group.PdfCount))

                Select Case mailType
                    Case MailProcessingType.DidiInvoiceMail
                        ' 只合并/命名滴滴成员；其余 PDF 仍按各自类型独立归档。
                        ' 否则一封"滴滴发票 + 无关增值税发票"的邮件会被合并成一个文件，
                        ' 第二张发票丢失独立可检索性与正确命名（静默合并两张会计凭证）。
                        Dim didiPdfs As New List(Of PdfAttachmentClassificationResult)()
                        Dim otherPdfs As New List(Of PdfAttachmentClassificationResult)()
                        For Each pc As PdfAttachmentClassificationResult In perPdf
                            If pc.IsDidi Then didiPdfs.Add(pc) Else otherPdfs.Add(pc)
                        Next
                        If didiPdfs.Count > 0 Then
                            FinishItem(batch, ProcessDidiMail(group, didiPdfs, reporter, processed), group, reporter, processed)
                        End If
                        ProcessPerPdfItems(batch, group, otherPdfs, reporter, processed)
                    Case MailProcessingType.GeneralInvoiceMail, MailProcessingType.MixedPdfMail, MailProcessingType.UnknownPdfOnlyMail
                        ProcessPerPdfItems(batch, group, perPdf, reporter, processed)
                    Case Else ' NoPdfMail（分组已过滤，一般不会到此）
                        ReportStage(reporter, processed, group, ArchiveStage.NoPdfSkipped)
                        FinishItem(batch, BuildSkippedItem(group), group, reporter, processed)
                End Select

            Catch ex As Exception
                Dim unk As New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex))
                AppLogger.Error("处理邮件异常：" & PrivacySafeFormatter.MaskSubject(group.MailSubject), ex)
                FinishItem(batch, BuildFailedItem(group, unk.UserMessage), group, reporter, processed)
            End Try

            processed += 1
        Next
        Return processed
    End Function

    ' ===== 滴滴合并归档 =====
    ''' <summary>
    ''' 合并归档一封邮件中的**滴滴**成员（发票 PDF + 行程单 PDF）为一个文件。
    ''' <paramref name="perPdf"/> 只应包含滴滴成员；非滴滴 PDF 由 ProcessPerPdfItems 独立处理。
    ''' </summary>
    Private Function ProcessDidiMail(group As MailPdfGroup, perPdf As List(Of PdfAttachmentClassificationResult),
                                     reporter As IArchiveProgressReporter, processed As Integer) As ArchiveItemResult
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = perPdf.Count, .ProcessingKind = "滴滴合并",
            .OriginalFileName = If(perPdf.Count > 0, perPdf(0).Attachment.OriginalFileName, Nothing)}
        For Each pc As PdfAttachmentClassificationResult In perPdf
            item.SourcePdfNames.Add(pc.Attachment.OriginalFileName)
            item.SourcePdfPaths.Add(pc.Attachment.TempFilePath)
        Next

        Dim recognized As New List(Of InvoiceRecognitionResult)()
        For Each pc As PdfAttachmentClassificationResult In perPdf
            recognized.Add(pc.Recognition)
        Next
        ReportStage(reporter, processed, group, ArchiveStage.ParsingFields)
        Dim mergedRecog As InvoiceRecognitionResult = _merger.Merge(recognized)
        item.DocumentType = mergedRecog.DocumentType
        item.RecognitionSource = mergedRecog.Source
        If mergedRecog.Invoice IsNot Nothing AndAlso mergedRecog.Invoice.Trips IsNot Nothing Then item.TripCount = mergedRecog.Invoice.Trips.Count

        ' 合并 PDF（发票在前、行程单在后）
        ReportStage(reporter, processed, group, ArchiveStage.Merging)
        Dim orderedPaths As List(Of String) = OrderDidiPaths(perPdf)
        Dim mergedTemp As String = PathHelper.GetNonConflictingPath(_tempDir, "merged_" & _runToken & "_" & group.Index & ".pdf")
        Dim mergeRes As Result = _mergeSvc.Merge(orderedPaths, mergedTemp)
        If Not mergeRes.IsSuccess Then
            item.Status = ProcessStatus.Failure
            Dim mergeErr As New AppError(AppErrorCode.PdfMergeFailed, mergeRes.Message) : mergeErr.LogSelf()
            item.Message = mergeErr.UserMessage
            Return item
        End If
        item.MergedTempPath = mergedTemp

        ReportStage(reporter, processed, group, ArchiveStage.Naming)
        Dim token As String = _runToken & "_" & group.Index.ToString()
        Dim plan As ArchiveTargetPlan = _planner.PlanTarget(_archiveFolder, mergedRecog.Invoice, mergedRecog.DocumentType,
                                                            item.OriginalFileName, token, group.MailSubject, group.PdfCount, mergedRecog.Source.ToString())
        ApplyNamePlan(item, plan)

        ReportStage(reporter, processed, group, ArchiveStage.Archiving)
        Dim exec As Result = _executor.Execute(mergedTemp, plan.FullPath)
        If exec.IsSuccess Then
            item.Status = MapRecognitionToProcess(mergedRecog.Status)
            item.Message = BuildItemReason(item.Status, mergedRecog)
            For Each pc As PdfAttachmentClassificationResult In perPdf
                SafeDeleteTemp(pc.Attachment.TempFilePath)
            Next
            SafeDeleteTemp(mergedTemp)
        Else
            item.Status = ProcessStatus.Failure
            ClearNamePlan(item)
            Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
            item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
        End If
        Return item
    End Function

    ''' <summary>逐个 PDF 独立归档（常规发票 / 未识别），单条失败不影响其它。</summary>
    Private Sub ProcessPerPdfItems(batch As ArchiveBatchResult, group As MailPdfGroup,
                                   pdfs As List(Of PdfAttachmentClassificationResult),
                                   reporter As IArchiveProgressReporter, processed As Integer)
        Dim idx As Integer = 0
        For Each pc As PdfAttachmentClassificationResult In pdfs
            ' 在每个 PDF 归档前响应取消，已归档的文件保留。
            If reporter IsNot Nothing AndAlso reporter.IsCancellationRequested Then
                _canceled = True
                Return
            End If
            idx += 1
            Dim it As ArchiveItemResult
            If pc.Classification = PdfAttachmentClassification.GeneralInvoicePdf Then
                ReportStage(reporter, processed, group, ArchiveStage.ProcessingGeneral)
                it = ProcessGeneralPdf(group, pc, idx)
            Else
                ReportStage(reporter, processed, group, ArchiveStage.ProcessingUnknown)
                it = ProcessUnknownPdf(group, pc, idx)
            End If
            FinishItem(batch, it, group, reporter, processed)
        Next
    End Sub

    ' ===== 常规发票（单 PDF 归档） =====
    Private Function ProcessGeneralPdf(group As MailPdfGroup, pc As PdfAttachmentClassificationResult, idx As Integer) As ArchiveItemResult
        Dim pdf As MailAttachmentItem = pc.Attachment
        Dim recog As InvoiceRecognitionResult = pc.Recognition
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = 1, .ProcessingKind = "常规发票", .OriginalFileName = pdf.OriginalFileName,
            .DocumentType = recog.DocumentType, .RecognitionSource = recog.Source}
        item.SourcePdfNames.Add(pdf.OriginalFileName)
        item.SourcePdfPaths.Add(pdf.TempFilePath)

        Try
            Dim token As String = _runToken & "_" & group.Index & "_" & idx
            Dim plan As ArchiveTargetPlan = _planner.PlanGeneralInvoiceTarget(_archiveFolder, recog.Invoice, pdf.OriginalFileName, token,
                                                                             group.MailSubject, recog.Source.ToString(), _generalTemplate)
            ApplyNamePlan(item, plan)
            Dim exec As Result = _executor.Execute(pdf.TempFilePath, plan.FullPath)
            If exec.IsSuccess Then
                item.Status = MapRecognitionToProcess(recog.Status)
                item.Message = "常规发票：" & BuildItemReason(item.Status, recog)
                SafeDeleteTemp(pdf.TempFilePath)
            Else
                item.Status = ProcessStatus.Failure
                ClearNamePlan(item)
                Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
                item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
            End If
        Catch ex As Exception
            item.Status = ProcessStatus.Failure
            item.Message = New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex)).UserMessage
            AppLogger.Error("处理常规发票异常：" & PrivacySafeFormatter.MaskFileName(pdf.OriginalFileName), ex)
        End Try
        Return item
    End Function

    ' ===== 未识别 PDF（未识别_原始文件名） =====
    Private Function ProcessUnknownPdf(group As MailPdfGroup, pc As PdfAttachmentClassificationResult, idx As Integer) As ArchiveItemResult
        Dim pdf As MailAttachmentItem = pc.Attachment
        Dim item As New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .SenderName = group.SenderName, .MailIndex = group.Index,
            .PdfCount = 1, .ProcessingKind = "未识别PDF", .OriginalFileName = pdf.OriginalFileName,
            .DocumentType = InvoiceDocumentType.Unknown, .RecognitionSource = RecognitionSource.None}
        item.SourcePdfNames.Add(pdf.OriginalFileName)
        item.SourcePdfPaths.Add(pdf.TempFilePath)

        Try
            Dim plan As ArchiveTargetPlan = _planner.PlanUnknownTarget(_archiveFolder, pdf.OriginalFileName)
            ApplyNamePlan(item, plan)
            Dim exec As Result = _executor.Execute(pdf.TempFilePath, plan.FullPath)
            If exec.IsSuccess Then
                item.Status = ProcessStatus.Success
                item.Message = "未识别 PDF，已按原文件名归档。"
                SafeDeleteTemp(pdf.TempFilePath)
            Else
                item.Status = ProcessStatus.Failure
                ClearNamePlan(item)
                Dim copyErr As New AppError(AppErrorCode.FileCopyFailed, exec.Message) : copyErr.LogSelf()
                item.Message = copyErr.UserMessage & "（临时文件已保留，可在本机临时目录排查）"
            End If
        Catch ex As Exception
            item.Status = ProcessStatus.Failure
            item.Message = New AppError(AppErrorCode.Unknown, ExceptionFormatter.ToUserMessage(ex)).UserMessage
            AppLogger.Error("处理未识别 PDF 异常：" & PrivacySafeFormatter.MaskFileName(pdf.OriginalFileName), ex)
        End Try
        Return item
    End Function

    Private Function BuildSkippedItem(group As MailPdfGroup) As ArchiveItemResult
        Return New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .MailIndex = group.Index, .PdfCount = group.PdfCount,
            .ProcessingKind = "跳过", .Status = ProcessStatus.Skipped, .Message = "无 PDF 附件，已跳过。"}
    End Function

    Private Function BuildFailedItem(group As MailPdfGroup, reason As String) As ArchiveItemResult
        Return New ArchiveItemResult With {
            .MailSubject = group.MailSubject, .MailIndex = group.Index, .PdfCount = group.PdfCount,
            .ProcessingKind = "失败", .Status = ProcessStatus.Failure, .Message = reason}
    End Function

    Private Sub ApplyNamePlan(item As ArchiveItemResult, plan As ArchiveTargetPlan)
        item.TargetPath = plan.FullPath
        item.FinalFileName = Path.GetFileName(plan.FullPath)
        item.NamingRule = plan.NamePlan.RuleName
        item.MissingFields = plan.NamePlan.MissingFields
        item.NamingFallbackTriggered = plan.NamePlan.FallbackTriggered
        item.UnknownPlaceholders = plan.NamePlan.UnknownPlaceholders
    End Sub

    Private Function BuildItemReason(status As ProcessStatus, recog As InvoiceRecognitionResult) As String
        Select Case status
            Case ProcessStatus.Success : Return "识别成功并归档。"
            Case ProcessStatus.PartialSuccess : Return If(String.IsNullOrEmpty(recog.Message), "部分成功，已归档（字段不完整）。", recog.Message)
            Case ProcessStatus.NeedsOcr : Return If(String.IsNullOrEmpty(recog.Message), UserFriendlyMessageProvider.Describe(AppErrorCode.LocalFieldsInsufficient).UserMessage, recog.Message)
            Case Else : Return If(String.IsNullOrEmpty(recog.Message), "已处理。", recog.Message)
        End Select
    End Function

    Private Sub FinishItem(batch As ArchiveBatchResult, item As ArchiveItemResult, group As MailPdfGroup, reporter As IArchiveProgressReporter, processed As Integer)
        ReportStage(reporter, processed, group, ArchiveStage.WritingResult)
        _carryNote = If(item.Status = ProcessStatus.Failure, "上一封邮件处理失败，已继续下一封。", Nothing)
        batch.Items.Add(item)
        AppLogger.Info(String.Format("邮件[{0}]【{1}】类型={2}, PDF={3}, 来源={4}, 状态={5}, 归档={6}, fallback={7}",
                                     item.MailIndex, PrivacySafeFormatter.MaskSubject(item.MailSubject), item.ProcessingKind,
                                     PrivacySafeFormatter.MaskFileList(item.SourcePdfNames),
                                     item.RecognitionSource, item.Status,
                                     If(String.IsNullOrEmpty(item.FinalFileName), "(未归档)", PrivacySafeFormatter.MaskFileName(item.FinalFileName)),
                                     item.NamingFallbackTriggered))
    End Sub

    ''' <summary>滴滴合并顺序：滴滴发票在前、滴滴行程单在后。仅接收滴滴成员。</summary>
    Private Function OrderDidiPaths(perPdf As List(Of PdfAttachmentClassificationResult)) As List(Of String)
        Dim paths As New List(Of String)()
        Dim order As PdfAttachmentClassification() = New PdfAttachmentClassification() {
            PdfAttachmentClassification.DidiInvoicePdf, PdfAttachmentClassification.DidiTripPdf}
        For Each k As PdfAttachmentClassification In order
            For Each pc As PdfAttachmentClassificationResult In perPdf
                If pc.Classification = k Then paths.Add(pc.Attachment.TempFilePath)
            Next
        Next
        Return paths
    End Function

    Private Sub ReportStage(reporter As IArchiveProgressReporter, processed As Integer, group As MailPdfGroup, stage As ArchiveStage)
        SafeReport(reporter, New ArchiveProgressInfo With {
            .TotalEmails = _total, .ProcessedEmails = processed,
            .CurrentEmailIndex = group.Index, .CurrentEmailSubject = group.MailSubject,
            .Stage = stage, .Note = _carryNote})
    End Sub

    Private Sub SafeReport(reporter As IArchiveProgressReporter, info As ArchiveProgressInfo)
        Try
            reporter.Report(info)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 识别状态 → 处理状态映射。
    ''' 注意：RecognitionStatus.Failure 必须映射为 ProcessStatus.Failure，否则硬失败会被
    ''' 记为"部分成功"且不计入 FailureCount，导致整批失败仍显示"部分成功"（失败对操作者不可见）。
    ''' </summary>
    Private Function MapRecognitionToProcess(status As RecognitionStatus) As ProcessStatus
        Select Case status
            Case RecognitionStatus.Success : Return ProcessStatus.Success
            Case RecognitionStatus.PartialSuccess : Return ProcessStatus.PartialSuccess
            Case RecognitionStatus.NeedsOcr : Return ProcessStatus.NeedsOcr
            Case RecognitionStatus.ConfigurationMissing : Return ProcessStatus.NeedsOcr
            Case RecognitionStatus.Failure : Return ProcessStatus.Failure
            Case Else : Return ProcessStatus.PartialSuccess
        End Select
    End Function

    Private Function ComputeOverall(batch As ArchiveBatchResult) As ProcessStatus
        If batch.TotalCount = 0 Then Return ProcessStatus.Skipped
        If batch.FailureCount = 0 AndAlso batch.NeedsOcrCount = 0 AndAlso batch.PartialCount = 0 Then Return ProcessStatus.Success
        If batch.SuccessCount = 0 AndAlso batch.PartialCount = 0 Then
            If batch.FailureCount = batch.TotalCount Then Return ProcessStatus.Failure
            Return ProcessStatus.PartialSuccess
        End If
        Return ProcessStatus.PartialSuccess
    End Function

    ''' <summary>读取设置；失败返回 Nothing，但必须留下日志痕迹（配置损坏不应完全不可见）。</summary>
    Private Function SafeSetting(getter As Func(Of String)) As String
        Try
            Return getter()
        Catch ex As Exception
            AppLogger.Warn("读取设置失败（按未配置处理）：" & ExceptionFormatter.ToUserMessage(ex))
            Return Nothing
        End Try
    End Function

    ''' <summary>删除本轮临时文件；失败记录日志（残留文件含票据信息，不应静默）。</summary>
    Private Sub SafeDeleteTemp(path As String)
        Try
            If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then File.Delete(path)
        Catch ex As Exception
            AppLogger.Warn("删除临时文件失败（将残留至保留期后自动清理）：" & PrivacySafeFormatter.MaskPath(path) &
                           " - " & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Sub

    ''' <summary>临时文件保留天数：超过该天数的历史临时文件在批次开始时清理，避免无限累积（含票据 PII）。</summary>
    Private Const TempRetentionDays As Integer = 7

    ''' <summary>清理临时目录中超过保留期的历史文件；失败不影响主流程。</summary>
    Private Sub PurgeOldTempFiles()
        Try
            If String.IsNullOrEmpty(_tempDir) OrElse Not Directory.Exists(_tempDir) Then Return
            Dim cutoff As DateTime = DateTime.UtcNow.AddDays(-TempRetentionDays)
            Dim removed As Integer = 0
            For Each f As String In Directory.GetFiles(_tempDir)
                Try
                    If File.GetLastWriteTimeUtc(f) < cutoff Then
                        File.Delete(f)
                        removed += 1
                    End If
                Catch
                    ' 单个文件被其它进程占用时跳过，下次再清。
                End Try
            Next
            If removed > 0 Then
                AppLogger.Info("已清理过期临时文件 " & removed.ToString(CultureInfo.InvariantCulture) &
                               " 个（保留期 " & TempRetentionDays.ToString(CultureInfo.InvariantCulture) & " 天）。")
            End If
        Catch ex As Exception
            AppLogger.Warn("清理过期临时文件失败（忽略）：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Sub

    ''' <summary>
    ''' 复制失败/未归档时清除已记录的目标路径与文件名，
    ''' 避免报告中出现实际并不存在的"归档路径"（两个条目显示同一路径而只有一个文件）。
    ''' </summary>
    Private Sub ClearNamePlan(item As ArchiveItemResult)
        item.TargetPath = Nothing
        item.FinalFileName = Nothing
    End Sub

End Class
