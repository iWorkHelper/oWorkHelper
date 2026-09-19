''' <summary>
''' 结构化应用错误：错误码 + 级别 + 用户提示 + 建议操作 + 仅供日志的详细信息。
''' 弹窗只展示 UserMessage/Suggestion；DetailForLog（含堆栈等）只写日志，绝不含 AK/SK/token。
''' </summary>
Public Class AppError

    Public Property Code As AppErrorCode
    Public Property Severity As ErrorSeverity
    ''' <summary>面向用户的简洁说明（“发生了什么”）。</summary>
    Public Property UserMessage As String
    ''' <summary>建议操作（“可以怎么做”）。</summary>
    Public Property Suggestion As String
    ''' <summary>仅供日志的详细信息（不显示给用户，不含敏感信息）。</summary>
    Public Property DetailForLog As String

    Public Sub New()
    End Sub

    Public Sub New(code As AppErrorCode, Optional detail As String = Nothing)
        Me.Code = code
        Dim info As AppError = UserFriendlyMessageProvider.Describe(code)
        Me.Severity = info.Severity
        Me.UserMessage = info.UserMessage
        Me.Suggestion = info.Suggestion
        Me.DetailForLog = detail
    End Sub

    ''' <summary>面向用户的完整提示（说明 + 建议）。</summary>
    Public Function ToUserText() As String
        If String.IsNullOrEmpty(Suggestion) Then
            Return UserMessage
        End If
        Return UserMessage & Environment.NewLine & "建议：" & Suggestion
    End Function

    ''' <summary>
    ''' 写入日志（级别按 Severity；含 DetailForLog，绝不含敏感信息）。
    ''' DetailForLog 常来自异常消息并内嵌本机路径（临时路径中含原始附件名），
    ''' 因此落盘前统一做路径脱敏。
    ''' </summary>
    Public Sub LogSelf()
        Dim detail As String = PrivacySafeFormatter.ScrubPaths(DetailForLog)
        Dim line As String = String.Format("[{0}] {1}{2}", Code, UserMessage,
                                            If(String.IsNullOrEmpty(detail), "", " | 详情: " & detail))
        Select Case Severity
            Case ErrorSeverity.Info : AppLogger.Info(line)
            Case ErrorSeverity.Warning : AppLogger.Warn(line)
            Case Else : AppLogger.Error(line)
        End Select
    End Sub

End Class
