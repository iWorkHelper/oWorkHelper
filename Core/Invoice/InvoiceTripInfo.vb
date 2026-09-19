''' <summary>
''' 网约车/滴滴行程单中的单条行程信息。
''' 一张行程单通常包含多条行程，故与 InvoiceInfo 为一对多关系。
''' </summary>
Public Class InvoiceTripInfo

    ''' <summary>乘车人。</summary>
    Public Property Passenger As String

    ''' <summary>出发时间（原始文本，保留原样，不强制解析为 DateTime）。</summary>
    Public Property DepartureTime As String

    ''' <summary>到达时间（原始文本）。</summary>
    Public Property ArrivalTime As String

    ''' <summary>起点。</summary>
    Public Property StartLocation As String

    ''' <summary>终点。</summary>
    Public Property EndLocation As String

    ''' <summary>车型/服务类型（如快车、专车）。</summary>
    Public Property ServiceType As String

    ''' <summary>订单号。</summary>
    Public Property OrderNumber As String

    ''' <summary>行程金额（原始文本）。</summary>
    Public Property TripAmount As String

    ''' <summary>城市。</summary>
    Public Property City As String

    ''' <summary>里程（公里，原始文本）。</summary>
    Public Property Mileage As String

    ''' <summary>行程序号（行程单中的序号，从 1 开始）。</summary>
    Public Property RowIndex As Integer

End Class
