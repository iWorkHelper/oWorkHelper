''' <summary>
''' 基于 ProgressForm 的进度上报器。所有异常内部吞掉，不影响归档主流程。
''' </summary>
Public Class UiArchiveProgressReporter
    Implements IArchiveProgressReporter

    Private ReadOnly _form As ProgressForm

    Public Sub New(form As ProgressForm)
        _form = form
    End Sub

    Public Sub Report(info As ArchiveProgressInfo) Implements IArchiveProgressReporter.Report
        Try
            If _form IsNot Nothing AndAlso Not _form.IsDisposed Then
                _form.UpdateProgress(info)
            End If
        Catch
            ' 忽略
        End Try
    End Sub

    ''' <summary>转发进度窗口的取消状态：用户在进度窗口点击“取消”后返回 True。</summary>
    Public ReadOnly Property IsCancellationRequested As Boolean Implements IArchiveProgressReporter.IsCancellationRequested
        Get
            Try
                Return _form IsNot Nothing AndAlso Not _form.IsDisposed AndAlso _form.CancelRequested
            Catch
                Return False
            End Try
        End Get
    End Property

End Class
