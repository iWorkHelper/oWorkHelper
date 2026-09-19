Imports System.IO
Imports System.Text

''' <summary>
''' 归档结果报告生成器：把本次批量结果写成用户可读文本报告，放在日志目录：
''' archive-report-yyyyMMdd-HHmmss.txt。报告不含 AK/SK/token 等敏感信息。
''' </summary>
Public Module ArchiveReportWriter

    ''' <summary>
    ''' 生成报告，返回文件路径；失败返回 Nothing（不影响主流程）。
    ''' </summary>
    ''' <param name="batch">批量结果。</param>
    ''' <param name="archiveFolder">归档目录（用于定位日志目录）。</param>
    ''' <param name="timestampToken">时间戳（yyyyMMdd-HHmmss，由调用方提供，避免脚本环境时钟限制）。</param>
    Public Function Write(batch As ArchiveBatchResult, archiveFolder As String, timestampToken As String) As String
        Try
            Dim logDir As String = PathHelper.GetLogDirectory(archiveFolder)
            If String.IsNullOrEmpty(logDir) Then Return Nothing
            Dim fileName As String = "archive-report-" & If(timestampToken, "unknown") & ".txt"
            Dim reportPath As String = Path.Combine(logDir, fileName)

            Dim sb As New StringBuilder()
            sb.AppendLine("iWorkHelper 归档结果报告")
            sb.AppendLine("批次 ID: " & If(batch.BatchId, "(无)"))
            sb.AppendLine("整体状态: " & batch.OverallStatus.ToString())
            sb.AppendLine(batch.BuildSummaryText())
            sb.AppendLine("需要配置 OCR 的邮件数: " & batch.NeedsOcrCount)
            sb.AppendLine(New String("="c, 60))

            Dim n As Integer = 0
            For Each it As ArchiveItemResult In batch.Items
                n += 1
                sb.AppendLine(String.Format("[{0}] 邮件《{1}》", it.MailIndex, PrivacySafeFormatter.MaskSubject(it.MailSubject)))
                sb.AppendLine("    状态: " & it.Status.ToString() & "  识别来源: " & it.RecognitionSource.ToString())
                sb.AppendLine("    PDF 数: " & it.PdfCount & "  原始附件: " & PrivacySafeFormatter.MaskFileList(it.SourcePdfNames))
                sb.AppendLine("    合并临时: " & If(String.IsNullOrEmpty(it.MergedTempPath), "-", PrivacySafeFormatter.MaskPath(it.MergedTempPath)))
                sb.AppendLine("    归档路径: " & If(String.IsNullOrEmpty(it.TargetPath), "(未归档)", PrivacySafeFormatter.MaskPath(it.TargetPath)))
                sb.AppendLine("    命名规则: " & If(it.NamingRule, "-") & "  fallback: " & it.NamingFallbackTriggered)
                If it.MissingFields IsNot Nothing AndAlso it.MissingFields.Count > 0 Then
                    sb.AppendLine("    缺失字段: " & JoinList(it.MissingFields))
                End If
                If it.UnknownPlaceholders IsNot Nothing AndAlso it.UnknownPlaceholders.Count > 0 Then
                    sb.AppendLine("    未知变量: " & JoinList(it.UnknownPlaceholders))
                End If
                If Not String.IsNullOrEmpty(it.Message) Then
                    ' 消息可能来自原始异常文本并内嵌本机临时路径（含原始附件名），
                    ' 报告又常写在共享归档目录，故必须脱敏后再落盘。
                    sb.AppendLine("    原因: " & PrivacySafeFormatter.ScrubPaths(it.Message))
                End If
            Next

            File.WriteAllText(reportPath, sb.ToString(), New UTF8Encoding(False))
            AppLogger.Info("已生成归档结果报告: " & PrivacySafeFormatter.MaskPath(reportPath))
            Return reportPath
        Catch ex As Exception
            AppLogger.Error("生成归档结果报告失败（忽略）。", ex)
            Return Nothing
        End Try
    End Function

    Private Function JoinList(items As System.Collections.Generic.List(Of String)) As String
        If items Is Nothing OrElse items.Count = 0 Then Return "-"
        Return String.Join(", ", items.ToArray())
    End Function

End Module
