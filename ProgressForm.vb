Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' 归档进度窗口。以“邮件”为单位显示总数/已处理/当前邮件/当前阶段/百分比。
''' 采用 Application.DoEvents 在关键阶段刷新（VSTO 归档在 UI 线程执行，避免后台线程访问 Outlook COM）。
'''
''' 取消：窗口提供“取消”按钮，点击后置位 <see cref="CancelRequested"/>。
''' 工作流在每个邮件/每个 PDF 的边界检查该标志并提前结束批次
''' （已归档的文件保留，未处理的邮件跳过），而不是强杀正在进行中的步骤。
''' </summary>
Public Class ProgressForm

    ''' <summary>用户是否已请求取消。</summary>
    Private _cancelRequested As Boolean = False

    ''' <summary>取消按钮（以代码创建，避免改动 Designer 文件）。</summary>
    Private WithEvents _btnCancel As Button

    ''' <summary>用户是否已请求取消（由工作流轮询）。</summary>
    Public ReadOnly Property CancelRequested As Boolean
        Get
            Return _cancelRequested
        End Get
    End Property

    Public Sub New()
        InitializeComponent()

        ' lblNote 在 Designer 中位于 y=135，超出原本 128 的客户区高度而被裁剪；
        ' 这里把窗口加高到能同时容纳提示文本与取消按钮。
        ClientSize = New Size(424, 172)

        _btnCancel = New Button()
        _btnCancel.Text = "取消"
        _btnCancel.Size = New Size(84, 26)
        _btnCancel.Location = New Point(ClientSize.Width - _btnCancel.Width - 12, 126)
        _btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        _btnCancel.TabIndex = 6
        Controls.Add(_btnCancel)
    End Sub

    ''' <summary>用进度信息刷新界面，并让 UI 有机会重绘（同时让取消按钮可被点击）。</summary>
    Public Sub UpdateProgress(info As ArchiveProgressInfo)
        If info Is Nothing Then Return
        Try
            lblTotal.Text = String.Format("共 {0} 封邮件，已处理 {1} 封", info.TotalEmails, info.ProcessedEmails)
            Dim subject As String = If(String.IsNullOrEmpty(info.CurrentEmailSubject), "-", info.CurrentEmailSubject)
            If subject.Length > 42 Then subject = subject.Substring(0, 42) & "…"
            lblCurrent.Text = String.Format("当前：第 {0}/{1} 封  {2}", info.CurrentEmailIndex, info.TotalEmails, subject)
            lblStage.Text = "阶段：" & info.StageText
            lblNote.Text = If(info.Note, "")
            progressBar1.Value = info.Percent
            Application.DoEvents()
        Catch
            ' 进度刷新失败不影响归档主流程。
        End Try
    End Sub

    ''' <summary>标记批次已因取消而结束，并禁用取消按钮。</summary>
    Public Sub MarkCanceled()
        Try
            _cancelRequested = True
            If _btnCancel IsNot Nothing Then
                _btnCancel.Enabled = False
                _btnCancel.Text = "已取消"
            End If
            lblNote.Text = "已取消：剩余邮件未处理，已归档的文件保留。"
            Application.DoEvents()
        Catch
        End Try
    End Sub

    Private Sub _btnCancel_Click(sender As Object, e As EventArgs) Handles _btnCancel.Click
        _cancelRequested = True
        ' 只置位标志：真正的中断发生在工作流的下一个检查点，
        ' 避免在 COM 调用/文件复制进行中强行打断而留下半成品。
        _btnCancel.Enabled = False
        _btnCancel.Text = "正在取消…"
        lblNote.Text = "正在取消，将在当前步骤结束后停止…"
        Application.DoEvents()
    End Sub

    Private Sub lblStage_Click(sender As Object, e As EventArgs) Handles lblStage.Click

    End Sub
End Class
