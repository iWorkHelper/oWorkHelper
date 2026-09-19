Imports System.Collections.Generic

''' <summary>
''' 常规发票表格区域检测：在逻辑行序列中定位「商品明细表格区」——
''' 从表头行（含 项目/名称/货物 等列名，且含 金额/税额）之后，到「合计 / 价税合计」行之前。
''' 找不到返回空区块。纯启发式，失败不抛异常。
''' </summary>
Public Module PdfTableRegionDetector

    Private Function Compact(ln As PdfTextLine) As String
        If ln Is Nothing OrElse ln.Text Is Nothing Then Return ""
        Return ln.Text.Replace(" ", "")
    End Function

    ''' <summary>检测商品明细区。</summary>
    Public Function DetectLineItemRegion(lines As List(Of PdfTextLine)) As PdfTextBlock
        Dim block As New PdfTextBlock()
        If lines Is Nothing OrElse lines.Count = 0 Then Return block

        Dim headerIdx As Integer = -1
        For i As Integer = 0 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            Dim hasNameCol As Boolean = c.Contains("项目名称") OrElse c.Contains("货物或应税劳务") OrElse c.Contains("货物或应税") OrElse (c.Contains("名称") AndAlso c.Contains("规格"))
            Dim hasMoneyCol As Boolean = c.Contains("金额") OrElse c.Contains("税额") OrElse c.Contains("税率")
            If hasNameCol AndAlso hasMoneyCol Then
                headerIdx = i
                Exit For
            End If
        Next
        If headerIdx < 0 Then Return block

        Dim endIdx As Integer = lines.Count
        For i As Integer = headerIdx + 1 To lines.Count - 1
            Dim c As String = Compact(lines(i))
            If c.Contains("合计") OrElse c.StartsWith("¥", StringComparison.Ordinal) OrElse c.StartsWith("￥", StringComparison.Ordinal) Then
                endIdx = i
                Exit For
            End If
        Next

        For i As Integer = headerIdx + 1 To endIdx - 1
            block.Lines.Add(lines(i))
        Next
        Return block
    End Function

End Module
