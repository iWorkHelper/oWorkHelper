Imports System.Collections.Generic

''' <summary>
''' 关键字段完整度评估。关键字段（阶段三暂定）：
'''  发票号码、开票日期、销售方名称、价税合计或金额、票据类型。
''' 缺失视为字段不足，但不丢弃已有字段。
''' </summary>
Public Module KeyFieldEvaluator

    ''' <summary>返回缺失的关键字段名列表。</summary>
    Public Function GetMissingKeyFields(inv As InvoiceInfo, docType As InvoiceDocumentType) As List(Of String)
        Dim missing As New List(Of String)()
        If inv Is Nothing Then
            missing.Add(InvoiceFieldNames.InvoiceNumber)
            missing.Add(InvoiceFieldNames.InvoiceDate)
            missing.Add(InvoiceFieldNames.SellerName)
            missing.Add(InvoiceFieldNames.TotalWithTax)
            missing.Add("票据类型")
            Return missing
        End If

        If String.IsNullOrWhiteSpace(inv.InvoiceNumber) Then missing.Add(InvoiceFieldNames.InvoiceNumber)
        If String.IsNullOrWhiteSpace(inv.InvoiceDate) Then missing.Add(InvoiceFieldNames.InvoiceDate)
        If String.IsNullOrWhiteSpace(inv.SellerName) Then missing.Add(InvoiceFieldNames.SellerName)
        If String.IsNullOrWhiteSpace(inv.TotalWithTax) AndAlso String.IsNullOrWhiteSpace(inv.Amount) Then
            missing.Add(InvoiceFieldNames.TotalWithTax)
        End If
        If docType = InvoiceDocumentType.Unknown Then missing.Add("票据类型")
        Return missing
    End Function

    ''' <summary>
    ''' 统一命名核心字段完整度：{乘车日期}{金额}{出发地点}{到达地点}。
    ''' 乘车日期=行程起止日期/首条行程出发时间/开票日期；金额=行程金额/价税合计/金额。
    ''' 行程类（或已解析出行程）还要求 出发地点(起点)+到达地点(终点)。
    ''' </summary>
    Public Function GetMissingNamingFields(inv As InvoiceInfo, docType As InvoiceDocumentType) As List(Of String)
        Dim missing As New List(Of String)()
        If inv Is Nothing Then
            missing.Add("乘车日期") : missing.Add("金额") : missing.Add("出发地点") : missing.Add("到达地点")
            Return missing
        End If

        Dim trip As InvoiceTripInfo = If(inv.Trips IsNot Nothing AndAlso inv.Trips.Count > 0, inv.Trips(0), Nothing)

        ' 乘车日期
        Dim hasRideDate As Boolean = (inv.ExtendedFields IsNot Nothing AndAlso inv.ExtendedFields.ContainsKey("行程起止日期") AndAlso Not String.IsNullOrWhiteSpace(inv.ExtendedFields("行程起止日期"))) _
            OrElse (trip IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(trip.DepartureTime)) _
            OrElse Not String.IsNullOrWhiteSpace(inv.InvoiceDate)
        If Not hasRideDate Then missing.Add("乘车日期")

        ' 金额
        Dim hasAmount As Boolean = (trip IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(trip.TripAmount)) _
            OrElse Not String.IsNullOrWhiteSpace(inv.TotalWithTax) OrElse Not String.IsNullOrWhiteSpace(inv.Amount)
        If Not hasAmount Then missing.Add("金额")

        ' 出发/到达地点：行程类或已有行程时要求
        Dim isRide As Boolean = docType = InvoiceDocumentType.RideTripStatement OrElse trip IsNot Nothing
        If isRide Then
            If trip Is Nothing OrElse String.IsNullOrWhiteSpace(trip.StartLocation) Then missing.Add("出发地点")
            If trip Is Nothing OrElse String.IsNullOrWhiteSpace(trip.EndLocation) Then missing.Add("到达地点")
        Else
            ' 非行程类（常规发票）：命名模板 {开票日期}_{金额}_{销售方名称} 需要销售方名称。
            ' 修复历史缺口：过去仅要求 日期+金额，导致缺销售方也判为充分而不回退在线 OCR。
            If docType = InvoiceDocumentType.VatInvoice AndAlso String.IsNullOrWhiteSpace(inv.SellerName) Then
                missing.Add(InvoiceFieldNames.SellerName)
            End If
        End If
        Return missing
    End Function

    ''' <summary>
    ''' 常规发票命名核心字段：开票日期 + 金额（价税合计/金额）+ 销售方名称。
    ''' 缺任一即字段不足（应回退在线 OCR）。
    ''' </summary>
    Public Function GetMissingGeneralInvoiceNamingFields(inv As InvoiceInfo) As List(Of String)
        Dim missing As New List(Of String)()
        If inv Is Nothing Then
            missing.Add(InvoiceFieldNames.InvoiceDate) : missing.Add("金额") : missing.Add(InvoiceFieldNames.SellerName)
            Return missing
        End If
        If String.IsNullOrWhiteSpace(inv.InvoiceDate) Then missing.Add(InvoiceFieldNames.InvoiceDate)
        If String.IsNullOrWhiteSpace(inv.TotalWithTax) AndAlso String.IsNullOrWhiteSpace(inv.Amount) Then missing.Add("金额")
        If String.IsNullOrWhiteSpace(inv.SellerName) Then missing.Add(InvoiceFieldNames.SellerName)
        Return missing
    End Function

    ''' <summary>统一命名核心字段是否充分。</summary>
    Public Function IsNamingSufficient(inv As InvoiceInfo, docType As InvoiceDocumentType) As Boolean
        Return GetMissingNamingFields(inv, docType).Count = 0
    End Function

End Module
