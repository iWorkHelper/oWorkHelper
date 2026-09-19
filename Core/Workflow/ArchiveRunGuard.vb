Imports System.Threading

''' <summary>
''' 归档任务的“唯一”运行锁（单一状态来源）。
'''
''' 设计要点：
'''  1. 全局只有这一处判断/设置“归档任务正在运行”，其它类不再各自维护运行标志。
'''  2. 使用 <see cref="Interlocked.CompareExchange"/> 原子获取，重复点击时只有第一次成功，杜绝竞争。
'''  3. 获取成功返回 <see cref="ArchiveRunToken"/>（IDisposable）；释放只能通过该 token。
'''  4. 纯内存、进程级：Outlook 进程重启后静态字段自动回到“未运行”，不使用磁盘锁/配置持久化，
'''     因此不存在“上次异常退出导致跨重启误判”的问题。
'''  5. 配合 Using / Try...Finally，可保证预检查失败、业务异常、进度窗口异常关闭等任何路径都能释放。
''' </summary>
Public NotInheritable Class ArchiveRunGuard

    Private Sub New()
    End Sub

    ' 0 = 空闲，1 = 运行中。所有读写均通过 Interlocked 原子完成。
    Private Shared _state As Integer = 0

    ' 持有者代号：每次成功获取自增。用于拒绝"非持有者"的释放请求
    ' （过期 token / 伪造 token 不得释放他人的运行锁）。
    Private Shared _generation As Long = 0

    ' 以下为“持有者”诊断信息，仅用于日志，不含任何敏感数据。
    Private Shared _batchId As String = Nothing
    Private Shared _threadId As Integer = 0
    Private Shared _acquiredAtTicks As Long = 0

    ''' <summary>
    ''' 尝试原子获取归档运行锁。
    ''' 成功返回非 Nothing 的 <see cref="ArchiveRunToken"/>；若已有任务在运行，返回 Nothing。
    ''' </summary>
    ''' <param name="batchId">可选批次标识，仅用于日志定位（不含敏感信息）。</param>
    Public Shared Function TryAcquire(Optional batchId As String = Nothing) As ArchiveRunToken
        ' 仅当当前为 0（空闲）时把它置为 1（运行中）；返回值为原值。
        If Interlocked.CompareExchange(_state, 1, 0) <> 0 Then
            Return Nothing ' 已被占用
        End If
        Dim ownerId As Long = Interlocked.Increment(_generation)
        _batchId = batchId
        _threadId = Thread.CurrentThread.ManagedThreadId
        _acquiredAtTicks = DateTime.UtcNow.Ticks
        Return New ArchiveRunToken(ownerId)
    End Function

    ''' <summary>当前是否有归档任务正在运行（原子读，不改变状态）。</summary>
    Public Shared ReadOnly Property IsRunning As Boolean
        Get
            Return Interlocked.CompareExchange(_state, 0, 0) = 1
        End Get
    End Property

    ''' <summary>
    ''' 描述当前运行锁持有者（用于获取失败时的日志）。未运行时返回占位。
    ''' 仅包含批次标识、线程 ID 与已持有毫秒数，绝不含敏感信息。
    ''' </summary>
    Public Shared Function DescribeHolder() As String
        If Not IsRunning Then Return "(无)"
        Dim heldMs As Long = 0
        Try
            heldMs = CLng((DateTime.UtcNow.Ticks - _acquiredAtTicks) \ TimeSpan.TicksPerMillisecond)
        Catch
        End Try
        Return String.Format("批次={0}, 持有线程={1}, 已持有={2}ms",
                             If(String.IsNullOrEmpty(_batchId), "(未命名)", _batchId), _threadId, heldMs)
    End Function

    ''' <summary>
    ''' 仅供 <see cref="ArchiveRunToken.Dispose"/> 调用：原子释放运行锁并清理持有者信息。
    ''' 只有持有者代号匹配时才释放；否则视为无效释放请求（no-op），
    ''' 避免过期/伪造的 token 释放掉他人正在使用的运行锁。
    ''' </summary>
    ''' <param name="ownerId">申请方持有者代号。</param>
    Friend Shared Sub Release(ownerId As Long)
        ' 锁未被释放前 _state 恒为 1，因此此处不可能与新的 TryAcquire 竞争成功。
        If Interlocked.Read(_generation) <> ownerId Then
            AppLogger.Warn("忽略无效的运行锁释放请求（非当前持有者）。")
            Return
        End If
        _batchId = Nothing
        _threadId = 0
        _acquiredAtTicks = 0
        Interlocked.Exchange(_state, 0)
    End Sub

End Class
