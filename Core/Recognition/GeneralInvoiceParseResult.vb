Imports System.Collections.Generic

''' <summary>
''' 常规发票本地解析的中间结果：全部字段候选（诊断用）、每字段择优结果、商品明细、发票类型与备注。
''' 由 GeneralInvoiceLocalRecognizer 组装，再落成统一的 InvoiceRecognitionResult。
''' </summary>
Public Class GeneralInvoiceParseResult

    Public Sub New()
        Candidates = New List(Of GeneralInvoiceFieldCandidate)()
        Chosen = New Dictionary(Of String, GeneralInvoiceFieldCandidate)()
        LineItems = New List(Of InvoiceLineItem)()
        Notes = New List(Of String)()
    End Sub

    ''' <summary>全部候选（未择优，供 --dump-candidates 诊断）。</summary>
    Public Property Candidates As List(Of GeneralInvoiceFieldCandidate)

    ''' <summary>每个字段名 → 择优后的候选。</summary>
    Public Property Chosen As Dictionary(Of String, GeneralInvoiceFieldCandidate)

    ''' <summary>商品明细。</summary>
    Public Property LineItems As List(Of InvoiceLineItem)

    ''' <summary>发票类型文案（增值税电子普通发票/专用发票/数电票/电子发票 等，尽力识别）。</summary>
    Public Property InvoiceTypeText As String

    ''' <summary>解析备注（如金额合计与价税合计不一致等，仅告警不失败）。</summary>
    Public Property Notes As List(Of String)

    ''' <summary>登记一个候选。</summary>
    Public Sub AddCandidate(c As GeneralInvoiceFieldCandidate)
        If c Is Nothing OrElse String.IsNullOrWhiteSpace(c.Value) Then Return
        Candidates.Add(c)
    End Sub

End Class
