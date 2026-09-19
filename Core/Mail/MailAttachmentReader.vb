Imports System.IO
Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook

''' <summary>
''' 批量读取 Outlook 当前选中邮件中的 PDF 附件，导出到临时工作目录。
''' 要求：
'''  - 支持批量选中；跳过非邮件项并记日志；只处理 PDF 附件；
'''  - 无 PDF 附件的邮件给出可读结果；临时文件名避免冲突、不覆盖用户文件；
'''  - 严格释放 Outlook COM 对象，避免泄漏。
''' </summary>
Public Class MailAttachmentReader

    Private Const PdfExtension As String = ".pdf"

    ''' <summary>
    ''' 从当前选择项读取 PDF 附件，并**按邮件分组**（每封邮件一组）。
    ''' 归档以邮件为单位：同组的多个 PDF 后续合并为一个 PDF。
    ''' </summary>
    Public Function ReadSelectedPdfGroups(application As Outlook.Application) As MailPdfGroupingResult
        Dim result As New MailPdfGroupingResult()

        If application Is Nothing Then
            result.AddMessage("无法获取 Outlook 应用对象。")
            AppLogger.Error("ReadSelectedPdfGroups: application 为空。")
            Return result
        End If

        Dim tempDir As String = PathHelper.GetTempWorkDirectory()
        Dim explorer As Outlook.Explorer = Nothing
        Dim selection As Outlook.Selection = Nothing
        Try
            explorer = application.ActiveExplorer()
            If explorer Is Nothing Then
                result.AddMessage("没有活动的 Outlook 窗口。")
                Return result
            End If
            selection = explorer.Selection
            If selection Is Nothing OrElse selection.Count = 0 Then
                result.AddMessage("未选中任何邮件。")
                AppLogger.Info("用户未选中任何邮件。")
                Return result
            End If

            result.SelectedItemCount = selection.Count
            AppLogger.Info("选中项数量: " & selection.Count.ToString())

            For i As Integer = 1 To selection.Count
                Dim item As Object = Nothing
                Try
                    item = selection.Item(i)
                    Dim mail As Outlook.MailItem = TryCast(item, Outlook.MailItem)
                    If mail Is Nothing Then
                        result.SkippedNonMailCount += 1
                        result.AddMessage("跳过第 " & i.ToString() & " 项：非邮件项。")
                        AppLogger.Info("跳过非邮件项，索引 " & i.ToString())
                        Continue For
                    End If

                    result.MailItemCount += 1
                    Dim group As New MailPdfGroup With {
                        .Index = i,
                        .MailSubject = SafeSubject(mail),
                        .SenderName = SafeSenderName(mail),
                        .ReceivedTime = SafeReceivedTime(mail)
                    }
                    ExportPdfAttachmentsToGroup(mail, tempDir, group, result)

                    If group.HasPdf Then
                        result.Groups.Add(group)
                        AppLogger.Info(String.Format("邮件【{0}】PDF 附件数={1}",
                                                     PrivacySafeFormatter.MaskSubject(group.MailSubject), group.PdfCount))
                    Else
                        result.MailsWithoutPdfCount += 1
                        result.AddMessage("邮件【" & group.MailSubject & "】无 PDF 附件。")
                    End If

                Catch exItem As Exception
                    result.AddMessage("处理第 " & i.ToString() & " 项时出错：" & ExceptionFormatter.ToUserMessage(exItem))
                    AppLogger.Error("处理选中项异常，索引 " & i.ToString(), exItem)
                Finally
                    ReleaseCom(item)
                End Try
            Next

            AppLogger.Info(String.Format("邮件分组读取完成: 邮件组={0}, 导出PDF={1}, 无PDF邮件={2}, 跳过非邮件={3}",
                                         result.Groups.Count, result.TotalPdfCount,
                                         result.MailsWithoutPdfCount, result.SkippedNonMailCount))
            Return result

        Catch ex As Exception
            result.AddMessage("读取选中邮件失败：" & ExceptionFormatter.ToUserMessage(ex))
            AppLogger.Error("ReadSelectedPdfGroups 异常。", ex)
            Return result
        Finally
            ReleaseCom(selection)
            ReleaseCom(explorer)
        End Try
    End Function

    ''' <summary>导出单封邮件的 PDF 附件到分组。</summary>
    Private Sub ExportPdfAttachmentsToGroup(mail As Outlook.MailItem, tempDir As String, group As MailPdfGroup, result As MailPdfGroupingResult)
        Dim attachments As Outlook.Attachments = Nothing
        Try
            attachments = mail.Attachments
            If attachments Is Nothing OrElse attachments.Count = 0 Then
                Return
            End If

            For a As Integer = 1 To attachments.Count
                Dim att As Outlook.Attachment = Nothing
                Try
                    att = attachments.Item(a)
                    Dim fileName As String = att.FileName
                    If String.IsNullOrEmpty(fileName) OrElse Not fileName.ToLowerInvariant().EndsWith(PdfExtension, StringComparison.Ordinal) Then
                        Continue For
                    End If

                    ' 附件名来自邮件，不可信：必须先剥掉任何目录部分（防御 "..\..\x.pdf" 之类的写入逃逸），
                    ' 再过文件名净化（非法字符会使 SaveAsFile 失败，导致好 PDF 被无谓跳过）。
                    Dim safeFileName As String = FileNameSanitizer.BuildFileName(
                        FileNameSanitizer.SanitizeBaseName(Path.GetFileName(fileName), "attachment"), ".pdf")

                    Dim targetPath As String = PathHelper.GetNonConflictingPath(tempDir, safeFileName)
                    att.SaveAsFile(targetPath)

                    group.Pdfs.Add(New MailAttachmentItem With {
                        .MailSubject = group.MailSubject,
                        .SenderName = group.SenderName,
                        .ReceivedTime = group.ReceivedTime,
                        .OriginalFileName = fileName,
                        .TempFilePath = targetPath,
                        .SizeBytes = SafeAttachmentSize(att)})
                    AppLogger.Info("已导出 PDF 附件: " & PrivacySafeFormatter.MaskFileName(fileName) &
                                   " -> " & PrivacySafeFormatter.MaskPath(targetPath))

                Catch exAtt As Exception
                    result.AddMessage("导出附件失败：" & ExceptionFormatter.ToUserMessage(exAtt))
                    AppLogger.Error("导出附件异常。", exAtt)
                Finally
                    ReleaseCom(att)
                End Try
            Next
        Finally
            ReleaseCom(attachments)
        End Try
    End Sub

    Private Function SafeSubject(mail As Outlook.MailItem) As String
        Try
            Return If(mail.Subject, "(无主题)")
        Catch
            Return "(无主题)"
        End Try
    End Function

    Private Function SafeSenderName(mail As Outlook.MailItem) As String
        Try
            Return If(mail.SenderName, String.Empty)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Function SafeReceivedTime(mail As Outlook.MailItem) As String
        Try
            Return mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        Catch
            Return String.Empty
        End Try
    End Function

    Private Function SafeAttachmentSize(att As Outlook.Attachment) As Long
        Try
            Return CLng(att.Size)
        Catch
            Return 0L
        End Try
    End Function

    ''' <summary>安全释放 COM 对象（统一委托到 ComRelease，避免多处实现走样）。</summary>
    Private Sub ReleaseCom(obj As Object)
        ComRelease.Release(obj)
    End Sub

End Class
