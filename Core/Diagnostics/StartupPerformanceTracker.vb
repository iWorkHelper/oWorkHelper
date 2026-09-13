Imports System.Diagnostics

''' <summary>
''' 轻量级启动性能追踪（使用 Stopwatch，不依赖重对象）。
''' 在 Startup 和关键阶段记录耗时。
''' 所有异常被内部吞掉，不影响主流程。
''' </summary>
Public Module StartupPerformanceTracker

    Private ReadOnly _trackers As New Dictionary(Of String, Stopwatch)()
    Private ReadOnly _SyncRoot As New Object()

    ''' <summary>
    ''' 开始计时一个阶段。可嵌套或重复调用同名阶段（将覆盖之前的计时）。
    ''' </summary>
    Public Sub StartStage(stageName As String)
        Try
            SyncLock _SyncRoot
                If String.IsNullOrWhiteSpace(stageName) Then Return
                Dim sw As New Stopwatch()
                sw.Start()
                _trackers(stageName) = sw
            End SyncLock
        Catch
            ' 诊断失败不影响主流程
        End Try
    End Sub

    ''' <summary>
    ''' 停止计时并返回耗时（毫秒）。如果未启动该阶段，返回 0。
    ''' </summary>
    Public Function EndStage(stageName As String) As Long
        Try
            SyncLock _SyncRoot
                If String.IsNullOrWhiteSpace(stageName) Then Return 0
                Dim sw As Stopwatch = Nothing
                If Not _trackers.TryGetValue(stageName, sw) Then Return 0
                If sw Is Nothing Then Return 0
                sw.Stop()
                Return sw.ElapsedMilliseconds
            End SyncLock
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 记录单个阶段的耗时到日志。
    ''' </summary>
    Public Sub LogStageDuration(stageName As String)
        Try
            Dim elapsed = EndStage(stageName)
            If elapsed >= 0 Then
                AppLogger.Debug(stageName & " 耗时：" & elapsed & "ms")
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 获取所有记录的耗时（用于诊断报告）。不清空数据。
    ''' </summary>
    Public Function GetAllTrackedDurations() As Dictionary(Of String, Long)
        Dim result As New Dictionary(Of String, Long)()
        Try
            SyncLock _SyncRoot
                For Each kvp In _trackers
                    Dim sw = kvp.Value
                    If sw IsNot Nothing Then
                        result(kvp.Key) = sw.ElapsedMilliseconds
                    End If
                Next
            End SyncLock
        Catch
        End Try
        Return result
    End Function

    ''' <summary>
    ''' 清空所有计时数据。
    ''' </summary>
    Public Sub Reset()
        Try
            SyncLock _SyncRoot
                _trackers.Clear()
            End SyncLock
        Catch
        End Try
    End Sub

End Module
