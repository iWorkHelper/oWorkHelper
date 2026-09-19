Imports System.IO

''' <summary>
''' 归档执行器：把临时 PDF 复制到目标路径。
''' 采用复制（非移动），保留临时文件供失败排查；临时文件清理由工作流统一处理。
''' 单个文件失败仅影响该文件，不抛出以免中断批次。
''' 绝不覆盖已有文件（含并发场景）。
''' </summary>
Public Class ArchiveExecutor

    ''' <summary>同目录临时后缀：先写 .partial 再原子改名，避免在最终文件名下留截断文件。</summary>
    Private Const PartialSuffix As String = ".partial"

    ''' <summary>
    ''' 执行复制。目标路径应为已解决冲突的完整路径（ArchivePlanner 产出）。
    ''' 三步走：复制到 .partial → 校验长度 → 原子改名到最终名。
    ''' 这样即使进程被杀或断电，也不会在**最终文件名**下留下截断 PDF
    ''' （原实现直接写最终名，中断后下次运行会因 File.Exists 报"已存在"而永远无法修复）。
    ''' 二次防御：若目标已存在（并发/异常），拒绝覆盖并返回失败。
    ''' </summary>
    Public Function Execute(sourceTempPath As String, targetPath As String) As Result
        Dim partialPath As String = targetPath & PartialSuffix
        Try
            If String.IsNullOrWhiteSpace(sourceTempPath) OrElse Not File.Exists(sourceTempPath) Then
                Return Result.Fail("源临时文件不存在：" & PrivacySafeFormatter.MaskPath(sourceTempPath))
            End If

            If File.Exists(targetPath) Then
                Dim existingLength As Long = -1
                Try
                    existingLength = New FileInfo(targetPath).Length
                Catch
                End Try
                If existingLength = 0 Then
                    ' 0 字节文件不可能是有效 PDF，明确提示可修复方式，而不是笼统说"已存在"。
                    Return Result.Fail("目标位置已有一个 0 字节文件（疑似上次中断的残片），未覆盖；请删除该文件后重试：" &
                                       PrivacySafeFormatter.MaskPath(targetPath))
                End If
                ' 绝不覆盖已有文件。
                Return Result.Fail("目标文件已存在，已跳过以避免覆盖：" & PrivacySafeFormatter.MaskPath(targetPath))
            End If

            Dim targetDir As String = Path.GetDirectoryName(targetPath)
            If Not PathHelper.EnsureDirectory(targetDir) Then
                Return Result.Fail("无法创建目标目录：" & PrivacySafeFormatter.MaskPath(targetDir))
            End If

            ' 1) 清理可能残留的 .partial，再复制到同目录临时名。
            SafeDeletePartial(partialPath)
            File.Copy(sourceTempPath, partialPath, overwrite:=False)

            ' 2) 长度校验：短拷贝不得被当作成功。
            Dim sourceLength As Long = New FileInfo(sourceTempPath).Length
            Dim copiedLength As Long = New FileInfo(partialPath).Length
            If sourceLength <> copiedLength Then
                SafeDeletePartial(partialPath)
                Return Result.Fail(String.Format("归档复制长度不一致（源 {0} 字节 / 目标 {1} 字节），已放弃本次归档。",
                                                 sourceLength, copiedLength))
            End If

            ' 3) 原子改名到最终名。File.Move 在目标已存在时会抛异常，因此仍然不会覆盖。
            File.Move(partialPath, targetPath)

            AppLogger.Info("已归档: " & PrivacySafeFormatter.MaskPath(sourceTempPath) &
                           " -> " & PrivacySafeFormatter.MaskPath(targetPath))
            Return Result.Ok(targetPath)

        Catch ex As Exception
            SafeDeletePartial(partialPath)
            AppLogger.Error("归档复制失败: " & PrivacySafeFormatter.MaskPath(targetPath), ex)
            Return Result.Fail("归档失败：" & ExceptionFormatter.ToUserMessage(ex))
        End Try
    End Function

    ''' <summary>删除可能残留的 .partial 中间文件；失败忽略（下次归档会重试清理）。</summary>
    Private Sub SafeDeletePartial(path As String)
        Try
            If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then File.Delete(path)
        Catch
        End Try
    End Sub

End Class
