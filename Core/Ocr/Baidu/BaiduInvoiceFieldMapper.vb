Imports System.Collections.Generic

''' <summary>
''' 百度返回字段 → InvoiceInfo / InvoiceTripInfo 的集中映射。
''' 原则：
'''  - 映射集中在此类，不散落到业务流程；
'''  - 已知字段入强类型属性，未知字段一律进 ExtendedFields（保留百度原始字段名）；
'''  - 保留置信度到 InvoiceField；单字段解析失败不影响整体；
'''  - 未知票据类型不报错。
''' 字段名参考百度智能财务票据识别/增值税发票识别常见返回，
''' 真实字段以样例返回 JSON 校准（见 FIELD_EXTRACTION_SPEC.md）。
''' </summary>
Public Class BaiduInvoiceFieldMapper

    ''' <summary>百度增值税发票字段名 → 内部字段名。</summary>
    Private Shared ReadOnly VatMap As Dictionary(Of String, String) = BuildVatMap()

    Private Shared Function BuildVatMap() As Dictionary(Of String, String)
        Dim m As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        m("InvoiceNum") = InvoiceFieldNames.InvoiceNumber
        m("InvoiceCode") = InvoiceFieldNames.InvoiceCode
        m("InvoiceDate") = InvoiceFieldNames.InvoiceDate
        m("SellerName") = InvoiceFieldNames.SellerName
        m("SellerRegisterNum") = InvoiceFieldNames.SellerTaxId
        m("PurchaserName") = InvoiceFieldNames.BuyerName
        m("PurchaserRegisterNum") = InvoiceFieldNames.BuyerTaxId
        m("CommodityName") = InvoiceFieldNames.ItemName
        m("CommodityType") = InvoiceFieldNames.Specification
        m("CommodityUnit") = InvoiceFieldNames.Unit
        m("CommodityNum") = InvoiceFieldNames.Quantity
        m("CommodityPrice") = InvoiceFieldNames.UnitPrice
        m("CommodityTaxRate") = InvoiceFieldNames.TaxRate
        m("TotalAmount") = InvoiceFieldNames.Amount
        m("TotalTax") = InvoiceFieldNames.TaxAmount
        ' 百度实际拼写为 AmountInFiguers（价税合计小写），一并兼容 AmountInFigures。
        m("AmountInFiguers") = InvoiceFieldNames.TotalWithTax
        m("AmountInFigures") = InvoiceFieldNames.TotalWithTax
        m("Checkcode") = InvoiceFieldNames.CheckCode
        m("CheckCode") = InvoiceFieldNames.CheckCode
        m("NoteDrawer") = InvoiceFieldNames.Drawer
        m("Payee") = InvoiceFieldNames.Payee
        ' 真实返回中"复核人"字段名为 Checker（非 Reviewer），两者都映射到复核人。
        m("Checker") = InvoiceFieldNames.Reviewer
        m("Reviewer") = InvoiceFieldNames.Reviewer
        m("Remarks") = InvoiceFieldNames.Remark
        Return m
    End Function

    ''' <summary>
    ''' 把解析后的百度文档映射进识别结果。返回是否提取到任何字段。
    ''' </summary>
    Public Function Map(doc As BaiduParsedDocument, result As InvoiceRecognitionResult) As Boolean
        If doc Is Nothing OrElse Not doc.HasItems Then
            Return False
        End If

        Dim mappedAny As Boolean = False
        Dim chosenType As InvoiceDocumentType = InvoiceDocumentType.Unknown

        For Each item As BaiduInvoiceItem In doc.Items
            Dim itemType As InvoiceDocumentType = BaiduInvoiceTypeMapper.Map(item.TypeRaw)

            ' 文档类型：优先增值税发票，其次出行类。
            If chosenType = InvoiceDocumentType.Unknown OrElse itemType = InvoiceDocumentType.VatInvoice Then
                chosenType = itemType
            End If

            If itemType = InvoiceDocumentType.RideTripStatement Then
                If MapRideItem(item, result) Then mappedAny = True
            Else
                If MapVatItem(item, result) Then mappedAny = True
            End If
        Next

        result.DocumentType = chosenType
        Return mappedAny
    End Function

    Private Function MapVatItem(item As BaiduInvoiceItem, result As InvoiceRecognitionResult) As Boolean
        Dim inv As InvoiceInfo = result.Invoice
        Dim mapped As Boolean = False

        For Each raw As BaiduRawField In item.RawFields
            If String.IsNullOrEmpty(raw.Name) OrElse String.IsNullOrWhiteSpace(raw.Word) Then
                Continue For
            End If

            Try
                Dim internalName As String = Nothing
                If VatMap.TryGetValue(raw.Name, internalName) Then
                    If AssignVatField(inv, internalName, raw.Word) Then
                        mapped = True
                    End If
                    AddField(result, internalName, raw)
                Else
                    ' 未知字段 → 扩展字段（保留百度原始名）。
                    inv.SetExtendedField(raw.Name, raw.Word)
                    AddField(result, raw.Name, raw)
                    mapped = True
                End If
            Catch
                ' 单字段失败不影响其它字段。
            End Try
        Next

        ' 滴滴等旅客运输发票：行程信息以 Passeng* 字段内嵌在 vat_invoice.result 中，
        ' 据真实返回按 row 分组构建多条行程明细。
        BuildTripsFromPassengFields(item, inv)

        Return mapped
    End Function

    ''' <summary>
    ''' 从 vat_invoice 内嵌的 Passeng*/Transport* 字段构建行程明细（按 row 分组）。
    ''' 字段名依据真实百度返回校准：PassengName/PassengDate/PassengOrigin/
    ''' PassengDestination/PassengVehicleType（均带 row）。
    ''' </summary>
    Private Sub BuildTripsFromPassengFields(item As BaiduInvoiceItem, inv As InvoiceInfo)
        Try
            Dim byRow As New Dictionary(Of Integer, InvoiceTripInfo)()
            For Each raw As BaiduRawField In item.RawFields
                If String.IsNullOrEmpty(raw.Name) OrElse String.IsNullOrWhiteSpace(raw.Word) Then Continue For
                If Not raw.Name.StartsWith("Passeng", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim key As Integer = If(raw.RowIndex > 0, raw.RowIndex, 1)
                Dim trip As InvoiceTripInfo = Nothing
                If Not byRow.TryGetValue(key, trip) Then
                    trip = New InvoiceTripInfo With {.RowIndex = key}
                    byRow(key) = trip
                End If

                Dim v As String = raw.Word.Trim()
                Select Case raw.Name
                    Case "PassengName" : trip.Passenger = v
                    Case "PassengDate"
                        trip.DepartureTime = v
                        If inv.ExtendedFields Is Nothing OrElse Not inv.ExtendedFields.ContainsKey("行程起止日期") Then inv.SetExtendedField("行程起止日期", v)
                    Case "PassengOrigin" : trip.StartLocation = v
                    Case "PassengDestination" : trip.EndLocation = v
                    Case "PassengVehicleType" : trip.ServiceType = v
                    Case Else
                        ' 其它 Passeng*（如 PassengClass/PassengIdNum）保留到扩展字段（已在主循环登记）。
                End Select
            Next

            If byRow.Count = 0 Then Return
            Dim keys As New List(Of Integer)(byRow.Keys)
            keys.Sort()
            For Each k As Integer In keys
                Dim t As InvoiceTripInfo = byRow(k)
                If Not String.IsNullOrEmpty(t.Passenger) OrElse Not String.IsNullOrEmpty(t.StartLocation) _
                   OrElse Not String.IsNullOrEmpty(t.EndLocation) OrElse Not String.IsNullOrEmpty(t.DepartureTime) Then
                    inv.Trips.Add(t)
                End If
            Next
            If inv.StatedTripCount = 0 Then inv.StatedTripCount = inv.Trips.Count
        Catch ex As Exception
            AppLogger.Warn("从 Passeng 字段构建行程失败（忽略）：" & ex.Message)
        End Try
    End Sub

    ''' <summary>写入强类型属性（首个非空值优先，避免多行覆盖）。返回是否写入。</summary>
    Private Function AssignVatField(inv As InvoiceInfo, internalName As String, value As String) As Boolean
        Dim v As String = value.Trim()
        Select Case internalName
            Case InvoiceFieldNames.InvoiceNumber : Return SetIfEmpty(Function() inv.InvoiceNumber, Sub(x) inv.InvoiceNumber = x, v)
            Case InvoiceFieldNames.InvoiceCode : Return SetIfEmpty(Function() inv.InvoiceCode, Sub(x) inv.InvoiceCode = x, v)
            Case InvoiceFieldNames.InvoiceDate : Return SetIfEmpty(Function() inv.InvoiceDate, Sub(x) inv.InvoiceDate = x, v)
            Case InvoiceFieldNames.SellerName : Return SetIfEmpty(Function() inv.SellerName, Sub(x) inv.SellerName = x, v)
            Case InvoiceFieldNames.SellerTaxId : Return SetIfEmpty(Function() inv.SellerTaxId, Sub(x) inv.SellerTaxId = x, v)
            Case InvoiceFieldNames.BuyerName : Return SetIfEmpty(Function() inv.BuyerName, Sub(x) inv.BuyerName = x, v)
            Case InvoiceFieldNames.BuyerTaxId : Return SetIfEmpty(Function() inv.BuyerTaxId, Sub(x) inv.BuyerTaxId = x, v)
            Case InvoiceFieldNames.ItemName : Return SetIfEmpty(Function() inv.ItemName, Sub(x) inv.ItemName = x, v)
            Case InvoiceFieldNames.Specification : Return SetIfEmpty(Function() inv.Specification, Sub(x) inv.Specification = x, v)
            Case InvoiceFieldNames.Unit : Return SetIfEmpty(Function() inv.Unit, Sub(x) inv.Unit = x, v)
            Case InvoiceFieldNames.Quantity : Return SetIfEmpty(Function() inv.Quantity, Sub(x) inv.Quantity = x, v)
            Case InvoiceFieldNames.UnitPrice : Return SetIfEmpty(Function() inv.UnitPrice, Sub(x) inv.UnitPrice = x, v)
            Case InvoiceFieldNames.Amount : Return SetIfEmpty(Function() inv.Amount, Sub(x) inv.Amount = x, v)
            Case InvoiceFieldNames.TaxRate : Return SetIfEmpty(Function() inv.TaxRate, Sub(x) inv.TaxRate = x, v)
            Case InvoiceFieldNames.TaxAmount : Return SetIfEmpty(Function() inv.TaxAmount, Sub(x) inv.TaxAmount = x, v)
            Case InvoiceFieldNames.TotalWithTax : Return SetIfEmpty(Function() inv.TotalWithTax, Sub(x) inv.TotalWithTax = x, v)
            Case InvoiceFieldNames.CheckCode : Return SetIfEmpty(Function() inv.CheckCode, Sub(x) inv.CheckCode = x, v)
            Case InvoiceFieldNames.Drawer : Return SetIfEmpty(Function() inv.Drawer, Sub(x) inv.Drawer = x, v)
            Case InvoiceFieldNames.Payee : Return SetIfEmpty(Function() inv.Payee, Sub(x) inv.Payee = x, v)
            Case InvoiceFieldNames.Reviewer : Return SetIfEmpty(Function() inv.Reviewer, Sub(x) inv.Reviewer = x, v)
            Case InvoiceFieldNames.Remark : Return SetIfEmpty(Function() inv.Remark, Sub(x) inv.Remark = x, v)
            Case Else
                inv.SetExtendedField(internalName, v)
                Return True
        End Select
    End Function

    ''' <summary>
    ''' 出行/行程类映射：百度字段名不稳定，统一保留到扩展字段；
    ''' **按 RowIndex 分组支持多条行程明细**（若百度以数组返回多行）。
    ''' 关键金额/日期回填到发票级字段，供命名使用。单字段失败不影响整体。
    ''' </summary>
    Private Function MapRideItem(item As BaiduInvoiceItem, result As InvoiceRecognitionResult) As Boolean
        Dim inv As InvoiceInfo = result.Invoice
        Dim mapped As Boolean = False

        ' 先把所有原始字段登记（扩展字段 + 诊断列表）。
        Dim byRow As New Dictionary(Of Integer, List(Of BaiduRawField))()
        For Each raw As BaiduRawField In item.RawFields
            If String.IsNullOrEmpty(raw.Name) OrElse String.IsNullOrWhiteSpace(raw.Word) Then
                Continue For
            End If
            inv.SetExtendedField(raw.Name, raw.Word.Trim())
            AddField(result, raw.Name, raw)
            mapped = True

            Dim key As Integer = raw.RowIndex
            If Not byRow.ContainsKey(key) Then byRow(key) = New List(Of BaiduRawField)()
            byRow(key).Add(raw)
        Next

        If Not mapped Then Return False

        ' 每个 RowIndex 构建一条行程（RowIndex=0 为标量/表头，也并入一条）。
        Dim rowKeys As New List(Of Integer)(byRow.Keys)
        rowKeys.Sort()
        For Each rk As Integer In rowKeys
            Dim trip As New InvoiceTripInfo With {.RowIndex = rk + 1}
            For Each raw As BaiduRawField In byRow(rk)
                AssignRideField(trip, inv, raw.Name, raw.Word.Trim())
            Next
            ' 仅当该行有实质内容才加入。
            If HasTripContent(trip) Then
                inv.Trips.Add(trip)
            End If
        Next

        ' 若分组后无有效行，但有字段，至少建一条汇总行程。
        If inv.Trips.Count = 0 Then
            inv.Trips.Add(New InvoiceTripInfo With {.RowIndex = 1, .TripAmount = inv.TotalWithTax})
        End If

        Return mapped
    End Function

    ''' <summary>
    ''' 依字段名关键词把值归入行程字段。用于独立 taxi_online_ticket/taxi_receipt（真实字段名未知，
    ''' 采用关键词最优匹配，兼容 Passeng*/英文/中文命名）。判定顺序有意先具体后宽泛。
    ''' </summary>
    Private Sub AssignRideField(trip As InvoiceTripInfo, inv As InvoiceInfo, name As String, v As String)
        Dim lower As String = name.ToLowerInvariant()
        If lower.Contains("arriv") OrElse lower.Contains("到达") Then
            If String.IsNullOrEmpty(trip.ArrivalTime) Then trip.ArrivalTime = v
        ElseIf lower.Contains("total") OrElse lower.Contains("amount") OrElse lower.Contains("fare") OrElse lower.Contains("金额") Then
            If String.IsNullOrEmpty(trip.TripAmount) Then trip.TripAmount = v
            If String.IsNullOrEmpty(inv.TotalWithTax) Then inv.TotalWithTax = v
        ElseIf lower.Contains("origin") OrElse lower.Contains("start") OrElse lower.Contains("from") OrElse lower.Contains("起") Then
            If String.IsNullOrEmpty(trip.StartLocation) Then trip.StartLocation = v
        ElseIf lower.Contains("dest") OrElse lower.Contains("终") OrElse lower.Contains("到点") OrElse lower.EndsWith("end", StringComparison.Ordinal) Then
            If String.IsNullOrEmpty(trip.EndLocation) Then trip.EndLocation = v
        ElseIf lower.Contains("mile") OrElse lower.Contains("里程") OrElse lower.Contains("distance") Then
            If String.IsNullOrEmpty(trip.Mileage) Then trip.Mileage = v
        ElseIf lower.Contains("vehicle") OrElse lower.Contains("车型") OrElse lower.Contains("service") OrElse lower.Contains("cartype") Then
            If String.IsNullOrEmpty(trip.ServiceType) Then trip.ServiceType = v
        ElseIf lower.Contains("order") OrElse lower.Contains("订单") Then
            If String.IsNullOrEmpty(trip.OrderNumber) Then trip.OrderNumber = v
        ElseIf lower.Contains("city") OrElse lower.Contains("城市") Then
            If String.IsNullOrEmpty(trip.City) Then trip.City = v
        ElseIf lower.Contains("passeng") OrElse lower.Contains("乘车人") OrElse lower.Contains("name") Then
            If String.IsNullOrEmpty(trip.Passenger) Then trip.Passenger = v
        ElseIf lower.Contains("time") OrElse lower.Contains("date") OrElse lower.Contains("depart") OrElse lower.Contains("时间") Then
            If String.IsNullOrEmpty(trip.DepartureTime) Then trip.DepartureTime = v
            If String.IsNullOrEmpty(inv.InvoiceDate) Then inv.InvoiceDate = v
        End If
    End Sub

    Private Function HasTripContent(trip As InvoiceTripInfo) As Boolean
        Return Not String.IsNullOrEmpty(trip.TripAmount) OrElse Not String.IsNullOrEmpty(trip.StartLocation) _
            OrElse Not String.IsNullOrEmpty(trip.EndLocation) OrElse Not String.IsNullOrEmpty(trip.DepartureTime) _
            OrElse Not String.IsNullOrEmpty(trip.OrderNumber) OrElse Not String.IsNullOrEmpty(trip.ServiceType)
    End Function

    Private Sub AddField(result As InvoiceRecognitionResult, displayName As String, raw As BaiduRawField)
        Dim f As New InvoiceField(displayName, raw.Word, "BaiduOcr:" & raw.Name)
        If raw.HasProbability Then
            f.Confidence = raw.Probability
        End If
        result.Fields.Add(f)
    End Sub

    Private Function SetIfEmpty(getter As Func(Of String), setter As Action(Of String), value As String) As Boolean
        If String.IsNullOrWhiteSpace(getter()) AndAlso Not String.IsNullOrWhiteSpace(value) Then
            setter(value)
            Return True
        End If
        Return False
    End Function

End Class
