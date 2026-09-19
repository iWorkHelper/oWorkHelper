Imports System.IO

''' <summary>
''' 路径与目录相关的辅助方法：工作目录、日志目录、唯一文件名生成、目录确保等。
''' 所有方法尽量不抛异常，供上层安全调用。
''' </summary>
Public Module PathHelper

    Public Const AppFolderName As String = "iWorkHelper"

    ''' <summary>
    ''' 应用根数据目录：%AppData%\iWorkHelper。
    ''' </summary>
    Public Function GetAppDataRoot() As String
        Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Return Path.Combine(appData, AppFolderName)
    End Function

    ''' <summary>
    ''' 日志目录：固定为 %AppData%\iWorkHelper\logs（O-45）。
    ''' 不再优先使用“归档目录\logs”：归档目录可能是共享/网络目录，多个用户会争用同一个
    ''' yyyy-MM-dd.log 文件名，并发追加会静默失败。日志文件名与轮转由 AppLogger 负责，
    ''' 这里只保证给出“每个 Windows 用户独立”的目录。
    ''' </summary>
    ''' <param name="archiveFolder">保留参数仅为兼容既有调用点，不再参与日志目录选择。</param>
    Public Function GetLogDirectory(Optional archiveFolder As String = Nothing) As String
        Dim logDir As String = Path.Combine(GetAppDataRoot(), "logs")
        EnsureDirectory(logDir)
        Return logDir
    End Function

    ''' <summary>
    ''' 临时工作目录：%AppData%\iWorkHelper\temp。用于暂存导出的 PDF 附件。
    ''' </summary>
    Public Function GetTempWorkDirectory() As String
        Dim tempDir As String = Path.Combine(GetAppDataRoot(), "temp")
        EnsureDirectory(tempDir)
        Return tempDir
    End Function

    ''' <summary>
    ''' 确保目录存在。返回是否可用。不抛异常。
    ''' </summary>
    Public Function EnsureDirectory(dirPath As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(dirPath) Then
                Return False
            End If
            If Not Directory.Exists(dirPath) Then
                Directory.CreateDirectory(dirPath)
            End If
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 在目标目录下为期望文件名生成不冲突的完整路径。
    ''' 若已存在同名文件，则自动追加 (1)、(2) ... 序号，绝不覆盖已有文件。
    ''' </summary>
    ''' <param name="targetDirectory">目标目录。</param>
    ''' <param name="desiredFileName">期望文件名（含扩展名）。</param>
    Public Function GetNonConflictingPath(targetDirectory As String, desiredFileName As String) As String
        Dim nameOnly As String = Path.GetFileNameWithoutExtension(desiredFileName)
        Dim ext As String = Path.GetExtension(desiredFileName)

        Dim candidate As String = Path.Combine(targetDirectory, desiredFileName)
        Dim index As Integer = 1
        While File.Exists(candidate)
            Dim newName As String = String.Format("{0}({1}){2}", nameOnly, index, ext)
            candidate = Path.Combine(targetDirectory, newName)
            index += 1
        End While
        Return candidate
    End Function

End Module
