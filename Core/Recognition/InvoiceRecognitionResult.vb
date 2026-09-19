Imports System.Collections.Generic

''' <summary>
''' 识别状态。必须能表达：成功、失败、部分成功、需要 OCR、配置缺失。
''' </summary>
Public Enum RecognitionStatus
    ''' <summary>成功（关键字段基本齐全）。</summary>
    Success = 0
    ''' <summary>失败（无法识别或异常）。</summary>
    Failure = 1
    ''' <summary>部分成功（提取到部分字段，但不完整）。</summary>
    PartialSuccess = 2
    ''' <summary>需要 OCR（本地文本层无效/疑似图片型，需在线 OCR）。</summary>
    NeedsOcr = 3
    ''' <summary>配置缺失（如在线 OCR 未启用/未配置）。</summary>
    ConfigurationMissing = 4
End Enum

''' <summary>
''' 识别来源：最终结果由哪条路径产生。
''' </summary>
Public Enum RecognitionSource
    ''' <summary>无（未识别/未开始）。</summary>
    None = 0
    ''' <summary>本地文本解析。</summary>
    LocalText = 1
    ''' <summary>百度在线 OCR。</summary>
    BaiduOcr = 2
    ''' <summary>本地 + 在线混合（如本地部分字段 + OCR 补充）。</summary>
    Mixed = 3
    ''' <summary>识别失败。</summary>
    Failed = 4
End Enum

''' <summary>
''' 票据识别结果。承载识别状态、文档类型、字段集合与结构化发票信息。
''' </summary>
Public Class InvoiceRecognitionResult

    Public Sub New()
        Fields = New List(Of InvoiceField)()
        Invoice = New InvoiceInfo()
        Messages = New List(Of String)()
        MissingKeyFields = New List(Of String)()
        Source = RecognitionSource.None
    End Sub

    Public Property Status As RecognitionStatus
    Public Property DocumentType As InvoiceDocumentType
    Public Property Message As String

    ''' <summary>识别来源，标记最终结果由哪种路径产生。</summary>
    Public Property Source As RecognitionSource

    ''' <summary>缺失的关键字段列表（供命名/诊断参考）。</summary>
    Public Property MissingKeyFields As List(Of String)

    ''' <summary>识别器名称（LocalTextInvoiceRecognizer / BaiduOcrInvoiceRecognizer）。</summary>
    Public Property RecognizerName As String

    ''' <summary>扁平字段列表（便于诊断）。</summary>
    Public Property Fields As List(Of InvoiceField)

    ''' <summary>结构化发票信息（供命名/归档使用）。</summary>
    Public Property Invoice As InvoiceInfo

    ''' <summary>原始文本（本地抽取或 OCR 返回，便于回溯）。</summary>
    Public Property RawText As String

    Public Property Messages As List(Of String)

    Public ReadOnly Property IsUsable As Boolean
        Get
            Return Status = RecognitionStatus.Success OrElse Status = RecognitionStatus.PartialSuccess
        End Get
    End Property

    Public Shared Function NeedsOcr(message As String, Optional rawText As String = "") As InvoiceRecognitionResult
        Return New InvoiceRecognitionResult With {
            .Status = RecognitionStatus.NeedsOcr,
            .Message = message,
            .RawText = rawText,
            .DocumentType = InvoiceDocumentType.Unknown
        }
    End Function

    Public Shared Function Failure(message As String) As InvoiceRecognitionResult
        Return New InvoiceRecognitionResult With {
            .Status = RecognitionStatus.Failure,
            .Message = message
        }
    End Function

End Class
