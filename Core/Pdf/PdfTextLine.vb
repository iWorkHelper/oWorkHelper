Imports System.Collections.Generic

''' <summary>
''' PDF 中的一个词（带坐标）。X=左边界，Right=右边界，Y=下边界（PDF 坐标系，Y 越大越靠上）。
''' </summary>
Public Class PdfTextWord
    Public Property Text As String
    Public Property X As Double
    Public Property Right As Double
    Public Property Y As Double

    Public ReadOnly Property CenterX As Double
        Get
            Return (X + Right) / 2.0
        End Get
    End Property
End Class

''' <summary>
''' 由坐标聚类重建出的一“行”：同一 Y 带内的词，按 X 从左到右排序。
''' 用于表格类版式（如滴滴行程单）的按列取单元格，避免依赖 page.Text 的拼接顺序。
''' </summary>
Public Class PdfTextLine
    Public Sub New()
        Words = New List(Of PdfTextWord)()
    End Sub

    ''' <summary>该行代表 Y（取行内词的中位/首个 Y）。</summary>
    Public Property Y As Double
    ''' <summary>所在页码（从 1 开始）。</summary>
    Public Property PageIndex As Integer
    ''' <summary>行内词（已按 X 升序）。</summary>
    Public Property Words As List(Of PdfTextWord)

    ''' <summary>行文本（词间以空格连接）。</summary>
    Public ReadOnly Property Text As String
        Get
            Dim parts As New List(Of String)()
            For Each w As PdfTextWord In Words
                If Not String.IsNullOrEmpty(w.Text) Then parts.Add(w.Text)
            Next
            Return String.Join(" ", parts.ToArray())
        End Get
    End Property

    ''' <summary>返回文本等于指定值的首个词（用于定位表头列）。</summary>
    Public Function FindWord(text As String) As PdfTextWord
        For Each w As PdfTextWord In Words
            If String.Equals(w.Text, text, StringComparison.Ordinal) Then Return w
        Next
        Return Nothing
    End Function
End Class
