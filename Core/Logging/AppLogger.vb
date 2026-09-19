Imports System.IO
Imports System.Globalization
Imports System.Text

''' <summary>
''' 日志级别。
''' </summary>
Public Enum LogLevel
    Debug_ = 0
    Info = 1
    Warn = 2
    [Error] = 3
End Enum

''' <summary>
''' 轻量级本地文件日志。要求：
'''  - 所有核心流程可调用，写入本地文本文件（默认 %AppData%\iWorkHelper\logs\yyyy-MM-dd.log）；
'''  - 任何写日志失败都不得导致主流程崩溃（内部全部吞掉异常）。
''' 非高并发场景，使用简单锁保证线程安全。
''' </summary>
Public Module AppLogger

    Private ReadOnly SyncRoot As New Object()
    Private _logDirectory As String = Nothing
    Private _initialized As Boolean = False

    ''' <summary>
    ''' 初始化日志目录。可传入归档目录，优先在其下建 logs 子目录。
    ''' 允许重复调用；失败时回退到 AppData。
    ''' </summary>
    Public Sub Initialize(Optional archiveFolder As String = Nothing)
        SyncLock SyncRoot
            Try
                _logDirectory = PathHelper.GetLogDirectory(archiveFolder)
            Catch
                _logDirectory = Nothing
            End Try
            _initialized = True
        End SyncLock
    End Sub

    ''' <summary>当前日志文件完整路径（供 UI 提示定位）。目录不可用时返回目录名或占位。</summary>
    Public Function CurrentLogPath() As String
        Try
            SyncLock SyncRoot
                If Not _initialized Then Initialize()
                If String.IsNullOrEmpty(_logDirectory) Then Return "(日志目录不可用)"
                Return Path.Combine(_logDirectory, LogFileName(DateTime.Now) & ".log")
            End SyncLock
        Catch
            Return "(日志目录不可用)"
        End Try
    End Function

    ''' <summary>
    ''' 日志文件名（不变式格式）。必须显式用 InvariantCulture：
    ''' 在非公历区域（如 th-TH / ar-SA）下 "yyyy-MM-dd" 会按该区域历法渲染年份，产生意外文件名。
    ''' </summary>
    Private Function LogFileName(now As DateTime) As String
        Return now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
    End Function

    Public Sub Info(message As String)
        Write(LogLevel.Info, message, Nothing)
    End Sub

    Public Sub Warn(message As String)
        Write(LogLevel.Warn, message, Nothing)
    End Sub

    Public Sub [Error](message As String, Optional ex As Exception = Nothing)
        Write(LogLevel.Error, message, ex)
    End Sub

    Public Sub Debug(message As String)
        Write(LogLevel.Debug_, message, Nothing)
    End Sub

    Private Sub Write(level As LogLevel, message As String, ex As Exception)
        Try
            SyncLock SyncRoot
                If Not _initialized Then
                    Initialize()
                End If

                Dim dir As String = _logDirectory
                If String.IsNullOrEmpty(dir) Then
                    ' 目录不可用时静默放弃，绝不抛出。
                    Return
                End If

                Dim fileName As String = LogFileName(DateTime.Now) & ".log"
                Dim fullPath As String = Path.Combine(dir, fileName)
                RollIfNeeded(fullPath, fileName)

                Dim sb As New StringBuilder()
                sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture))
                sb.Append(" [")
                sb.Append(LevelText(level))
                sb.Append("] ")
                sb.Append(If(message, String.Empty))
                If ex IsNot Nothing Then
                    sb.AppendLine()
                    sb.Append(ExceptionFormatter.ToLogText(ex))
                End If
                sb.AppendLine()

                File.AppendAllText(fullPath, sb.ToString(), Encoding.UTF8)
            End SyncLock
        Catch
            ' 日志绝不影响主流程。
        End Try
    End Sub

    ''' <summary>单个日志文件大小上限：超出后滚动，避免单日无限增长。</summary>
    Private Const MaxLogBytes As Long = 5 * 1024 * 1024

    ''' <summary>保留的滚动历史文件个数（.1 ... .N）。</summary>
    Private Const MaxRolledFiles As Integer = 5

    ''' <summary>
    ''' 单文件超过上限时滚动为 name.log.1、name.log.2…，并只保留最近 MaxRolledFiles 个。
    ''' 滚动失败不影响本次写入（继续追加到原文件）。
    ''' </summary>
    Private Sub RollIfNeeded(fullPath As String, baseName As String)
        Try
            If Not File.Exists(fullPath) Then Return
            If New FileInfo(fullPath).Length < MaxLogBytes Then Return

            Dim dir As String = Path.GetDirectoryName(fullPath)
            For i As Integer = MaxRolledFiles - 1 To 1 Step -1
                Dim src As String = RolledPath(dir, baseName, i)
                Dim dst As String = RolledPath(dir, baseName, i + 1)
                If File.Exists(src) Then
                    If File.Exists(dst) Then File.Delete(dst)
                    File.Move(src, dst)
                End If
            Next
            Dim first As String = RolledPath(dir, baseName, 1)
            If File.Exists(first) Then File.Delete(first)
            File.Move(fullPath, first)
        Catch
            ' 滚动失败时继续追加写原文件。
        End Try
    End Sub

    Private Function RolledPath(dir As String, baseName As String, index As Integer) As String
        Return Path.Combine(dir, baseName & "." & index.ToString(CultureInfo.InvariantCulture))
    End Function

    Private Function LevelText(level As LogLevel) As String
        Select Case level
            Case LogLevel.Debug_
                Return "DEBUG"
            Case LogLevel.Info
                Return "INFO "
            Case LogLevel.Warn
                Return "WARN "
            Case LogLevel.Error
                Return "ERROR"
            Case Else
                Return "INFO "
        End Select
    End Function

End Module
