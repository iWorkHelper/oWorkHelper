Imports System.Collections.Generic

''' <summary>
''' 统一处理状态枚举，供各流程复用。
''' </summary>
Public Enum ProcessStatus
    ''' <summary>成功。</summary>
    Success = 0
    ''' <summary>失败。</summary>
    Failure = 1
    ''' <summary>部分成功（批次中部分项失败）。</summary>
    PartialSuccess = 2
    ''' <summary>跳过（不满足处理条件，如非 PDF、无附件）。</summary>
    Skipped = 3
    ''' <summary>需要 OCR（本地文本层无效，疑似图片型）。</summary>
    NeedsOcr = 4
    ''' <summary>配置缺失（如在线 OCR 未配置、归档目录未设置）。</summary>
    ConfigurationMissing = 5
End Enum

''' <summary>
''' 统一结果对象（无返回值）。所有核心流程返回该类型，
''' 以避免仅靠异常/MessageBox 传递状态。
''' </summary>
Public Class Result

    Private ReadOnly _messages As New List(Of String)()

    Public Property Status As ProcessStatus
    Public Property Message As String

    ''' <summary>过程中收集到的明细信息（成功/警告/错误逐条记录）。</summary>
    Public ReadOnly Property Messages As List(Of String)
        Get
            Return _messages
        End Get
    End Property

    Public ReadOnly Property IsSuccess As Boolean
        Get
            Return Status = ProcessStatus.Success
        End Get
    End Property

    Public Sub AddMessage(text As String)
        If Not String.IsNullOrEmpty(text) Then
            _messages.Add(text)
        End If
    End Sub

    Public Shared Function Ok(Optional message As String = "") As Result
        Return New Result With {.Status = ProcessStatus.Success, .Message = message}
    End Function

    Public Shared Function Fail(message As String) As Result
        Return New Result With {.Status = ProcessStatus.Failure, .Message = message}
    End Function

    Public Shared Function ConfigMissing(message As String) As Result
        Return New Result With {.Status = ProcessStatus.ConfigurationMissing, .Message = message}
    End Function

End Class
