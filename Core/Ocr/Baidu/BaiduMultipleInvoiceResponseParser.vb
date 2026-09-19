Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports System.Web.Script.Serialization

''' <summary>
''' 百度原始字段（保留原始字段名、值、置信度、位置）。
''' </summary>
Public Class BaiduRawField
    Public Property Name As String
    Public Property Word As String
    Public Property HasProbability As Boolean
    Public Property Probability As Double
    Public Property Location As String
    Public Property RowIndex As Integer
End Class

''' <summary>
''' 单张票据（words_result 数组中的一项）。
''' </summary>
Public Class BaiduInvoiceItem
    Public Sub New()
        RawFields = New List(Of BaiduRawField)()
    End Sub
    Public Property TypeRaw As String
    Public Property RawFields As List(Of BaiduRawField)
End Class

''' <summary>
''' 解析后的整份文档。
''' </summary>
Public Class BaiduParsedDocument
    Public Sub New()
        Items = New List(Of BaiduInvoiceItem)()
    End Sub
    Public Property WordsResultNum As Integer
    Public Property PdfFileSize As String
    Public Property Items As List(Of BaiduInvoiceItem)
    Public Property ParseError As String

    Public ReadOnly Property HasItems As Boolean
        Get
            Return Items IsNot Nothing AndAlso Items.Count > 0
        End Get
    End Property
End Class

''' <summary>
''' 使用 JavaScriptSerializer 弱类型解析百度智能财务票据识别返回。
''' 不硬套强类型完整模型：result 结构随票据类型变化，逐字段泛化提取，
''' 保留原始字段名/值/置信度/位置，无法映射的字段留待 FieldMapper 处理。
''' </summary>
Public Class BaiduMultipleInvoiceResponseParser

    ''' <summary>
    ''' JSON 体上限（16 MB）。多票识别响应正常远小于此值；
    ''' 设上限而非 Integer.MaxValue，使超大/异常响应以可处理的解析失败结束，
    ''' 而不是在 DeserializeObject 中抛出 OutOfMemoryException。
    ''' </summary>
    Private Const MaxJsonBytes As Integer = 16 * 1024 * 1024

    Public Function Parse(rawJson As String) As BaiduParsedDocument
        Dim doc As New BaiduParsedDocument()
        Try
            If String.IsNullOrWhiteSpace(rawJson) Then
                doc.ParseError = "空响应"
                Return doc
            End If

            Dim serializer As New JavaScriptSerializer()
            ' 上限而非 Integer.MaxValue：避免恶意/异常的超大响应把进程内存吃满
            ' （Integer.MaxValue 在超大输入下会触发 OutOfMemory 而非可处理的解析失败）。
            serializer.MaxJsonLength = MaxJsonBytes
            Dim root As Dictionary(Of String, Object) = TryCast(serializer.DeserializeObject(rawJson), Dictionary(Of String, Object))
            If root Is Nothing Then
                doc.ParseError = "根节点无法解析"
                Return doc
            End If

            If root.ContainsKey("pdf_file_size") AndAlso root("pdf_file_size") IsNot Nothing Then
                doc.PdfFileSize = Convert.ToString(root("pdf_file_size"))
            End If
            If root.ContainsKey("words_result_num") AndAlso root("words_result_num") IsNot Nothing Then
                Dim n As Integer
                Integer.TryParse(Convert.ToString(root("words_result_num")), n)
                doc.WordsResultNum = n
            End If

            Dim wordsResult As Object = If(root.ContainsKey("words_result"), root("words_result"), Nothing)
            Dim arr As Object() = TryCast(wordsResult, Object())
            If arr Is Nothing Then
                ' 有些接口 words_result 可能是对象而非数组；容错处理。
                Dim single_ As Dictionary(Of String, Object) = TryCast(wordsResult, Dictionary(Of String, Object))
                If single_ IsNot Nothing Then
                    doc.Items.Add(ParseItem(single_))
                End If
                Return doc
            End If

            For Each element As Object In arr
                Dim itemMap As Dictionary(Of String, Object) = TryCast(element, Dictionary(Of String, Object))
                If itemMap IsNot Nothing Then
                    doc.Items.Add(ParseItem(itemMap))
                End If
            Next

            Return doc

        Catch ex As Exception
            doc.ParseError = ExceptionFormatter.ToUserMessage(ex)
            AppLogger.Error("解析百度 OCR 返回 JSON 异常。", ex)
            Return doc
        End Try
    End Function

    Private Function ParseItem(itemMap As Dictionary(Of String, Object)) As BaiduInvoiceItem
        Dim item As New BaiduInvoiceItem()
        If itemMap.ContainsKey("type") AndAlso itemMap("type") IsNot Nothing Then
            item.TypeRaw = Convert.ToString(itemMap("type"))
        End If

        ' result 可能是对象（字段字典），也可能直接把字段铺在 item 上。
        Dim resultObj As Object = If(itemMap.ContainsKey("result"), itemMap("result"), Nothing)
        Dim resultMap As Dictionary(Of String, Object) = TryCast(resultObj, Dictionary(Of String, Object))
        If resultMap Is Nothing Then
            ' 退而求其次：把 item 本身当字段容器（排除 type）。
            resultMap = New Dictionary(Of String, Object)()
            For Each kv In itemMap
                If Not String.Equals(kv.Key, "type", StringComparison.OrdinalIgnoreCase) Then
                    resultMap(kv.Key) = kv.Value
                End If
            Next
        End If

        For Each kv In resultMap
            ExtractField(kv.Key, kv.Value, item.RawFields)
        Next

        Return item
    End Function

    ''' <summary>
    ''' 把单个字段值泛化为一到多个 BaiduRawField。
    ''' 值可能是：字符串 / 对象{word,probability,location} / 数组[对象...]（多行）。
    ''' </summary>
    Private Sub ExtractField(name As String, value As Object, sink As List(Of BaiduRawField))
        If value Is Nothing Then
            Return
        End If

        Dim arr As Object() = TryCast(value, Object())
        If arr IsNot Nothing Then
            Dim rowIndex As Integer = 0
            For Each element As Object In arr
                AppendFieldFromElement(name, element, rowIndex, sink)
                rowIndex += 1
            Next
            Return
        End If

        Dim map As Dictionary(Of String, Object) = TryCast(value, Dictionary(Of String, Object))
        If map IsNot Nothing Then
            AppendFieldFromElement(name, map, 0, sink)
            Return
        End If

        ' 标量值
        sink.Add(New BaiduRawField With {.Name = name, .Word = Convert.ToString(value), .RowIndex = 0})
    End Sub

    Private Sub AppendFieldFromElement(name As String, element As Object, rowIndex As Integer, sink As List(Of BaiduRawField))
        Dim map As Dictionary(Of String, Object) = TryCast(element, Dictionary(Of String, Object))
        If map Is Nothing Then
            sink.Add(New BaiduRawField With {.Name = name, .Word = Convert.ToString(element), .RowIndex = rowIndex})
            Return
        End If

        Dim field As New BaiduRawField With {.Name = name, .RowIndex = rowIndex}
        If map.ContainsKey("word") AndAlso map("word") IsNot Nothing Then
            field.Word = Convert.ToString(map("word"))
        ElseIf map.ContainsKey("words") AndAlso map("words") IsNot Nothing Then
            field.Word = Convert.ToString(map("words"))
        Else
            field.Word = FlattenScalarDictionary(map)
        End If

        ' 行号：百度真实返回中明细字段带 "row"（如 "1"），据此定行以支持多条明细。
        If map.ContainsKey("row") AndAlso map("row") IsNot Nothing Then
            Dim rowNo As Integer
            If Integer.TryParse(Convert.ToString(map("row")), rowNo) Then
                field.RowIndex = rowNo
            End If
        End If

        ' 置信度：probability 可能是对象 {average,min,...} 或标量。
        ' 数值一律按 InvariantCulture 解析（JSON 数字固定用 "." 作小数点）。
        If map.ContainsKey("probability") AndAlso map("probability") IsNot Nothing Then
            Dim probMap As Dictionary(Of String, Object) = TryCast(map("probability"), Dictionary(Of String, Object))
            If probMap IsNot Nothing AndAlso probMap.ContainsKey("average") Then
                Dim p As Double
                If Double.TryParse(Convert.ToString(probMap("average"), CultureInfo.InvariantCulture),
                                   NumberStyles.Float, CultureInfo.InvariantCulture, p) Then
                    field.HasProbability = True
                    field.Probability = p
                End If
            Else
                Dim p As Double
                If Double.TryParse(Convert.ToString(map("probability"), CultureInfo.InvariantCulture),
                                   NumberStyles.Float, CultureInfo.InvariantCulture, p) Then
                    field.HasProbability = True
                    field.Probability = p
                End If
            End If
        End If

        ' 位置：location 对象序列化为紧凑文本保留。
        If map.ContainsKey("location") AndAlso map("location") IsNot Nothing Then
            field.Location = SerializeLocation(map("location"))
        End If

        sink.Add(field)
    End Sub

    Private Function FlattenScalarDictionary(map As Dictionary(Of String, Object)) As String
        Dim sb As New StringBuilder()
        For Each kv In map
            Dim v As String = Convert.ToString(kv.Value)
            If Not String.IsNullOrEmpty(v) Then
                If sb.Length > 0 Then sb.Append(" ")
                sb.Append(v)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function SerializeLocation(loc As Object) As String
        Dim map As Dictionary(Of String, Object) = TryCast(loc, Dictionary(Of String, Object))
        If map Is Nothing Then
            Return Convert.ToString(loc)
        End If
        Dim parts As New List(Of String)()
        For Each k As String In New String() {"left", "top", "width", "height"}
            If map.ContainsKey(k) Then
                parts.Add(k & "=" & Convert.ToString(map(k)))
            End If
        Next
        Return String.Join(",", parts.ToArray())
    End Function

End Class
