Imports System.Collections.Generic
Imports System.Text.RegularExpressions

''' <summary>
''' 命名模板渲染结果。
''' </summary>
Public Class NamingRenderResult
    Public Sub New()
        MissingPlaceholders = New List(Of String)()
        UnknownPlaceholders = New List(Of String)()
    End Sub
    ''' <summary>渲染出的文件名（已清理非法字符、补 .pdf）。渲染失败为 Nothing。</summary>
    Public Property FileName As String
    ''' <summary>模板中引用但值为空的已知占位符（被跳过）。</summary>
    Public Property MissingPlaceholders As List(Of String)
    ''' <summary>模板中出现的未知占位符（不认识，按空处理并提示）。</summary>
    Public Property UnknownPlaceholders As List(Of String)
    ''' <summary>
    ''' 字段严重不足：模板中至少有一个占位符，但**没有任何占位符产出值**（或渲染结果本身为空白）。
    ''' 判据是“字段充分性”而非“渲染结果是否空白”，避免 {金额}（{出发地点}）全空时产出「（）.pdf」这类垃圾文件名。
    ''' </summary>
    Public Property WasEmpty As Boolean
End Class

''' <summary>
''' 命名模板引擎：把 {占位符} 替换为字段值。
''' 规则：
'''  - 已知占位符值为空 → 跳过并记入 Missing，且避免产生连续分隔符；
'''  - 未知占位符 → 记入 Unknown（供日志/结果提示），按空处理；
'''  - 清理 Windows 非法字符；模板错误不抛出（由调用方回退默认规则）。
''' </summary>
Public Module NamingTemplateEngine

    Private ReadOnly PlaceholderPattern As New Regex("\{([^{}]+)\}")

    ''' <summary>
    ''' 渲染模板。values 应包含全部"已知"占位符键（值可为空）；不在其中的键视为未知。
    ''' </summary>
    Public Function Render(template As String, values As IDictionary(Of String, String), fallbackBaseName As String) As NamingRenderResult
        Dim result As New NamingRenderResult()

        Try
            If String.IsNullOrWhiteSpace(template) Then
                Return result ' FileName = Nothing → 调用方回退
            End If

            Dim missing As List(Of String) = result.MissingPlaceholders
            Dim unknown As List(Of String) = result.UnknownPlaceholders

            ' 记录占位符总数与“产出非空值”的占位符数：用于按字段充分性判定回退，
            ' 而不是看渲染结果是否空白（{金额}（{出发地点}）全空时会渲染出「（）」）。
            Dim placeholderCount As Integer = 0
            Dim valuedPlaceholderCount As Integer = 0

            Dim replaced As String = PlaceholderPattern.Replace(template, Function(m)
                                                                              placeholderCount += 1
                                                                              Dim key As String = m.Groups(1).Value.Trim()
                                                                              Dim v As String = Nothing
                                                                              If values IsNot Nothing AndAlso values.TryGetValue(key, v) Then
                                                                                  If String.IsNullOrWhiteSpace(v) Then
                                                                                      missing.Add(key)
                                                                                      Return String.Empty
                                                                                  End If
                                                                                  valuedPlaceholderCount += 1
                                                                                  Return v
                                                                              Else
                                                                                  unknown.Add(key)
                                                                                  Return String.Empty
                                                                              End If
                                                                          End Function)

            ' 折叠因空占位符产生的连续分隔符（下划线/连字符），并去除首尾分隔符。
            replaced = Regex.Replace(replaced, "[_\-]{2,}", "_")
            replaced = replaced.Trim(New Char() {"_"c, "-"c, "."c, " "c})

            ' 字段充分性：模板含占位符但无一产出值 → 字段严重不足（即使渲染结果非空白，如「（）」）。
            ' 纯字面量模板（无占位符）不按此判据回退，仍以渲染结果是否空白为准。
            Dim noFieldValue As Boolean = placeholderCount > 0 AndAlso valuedPlaceholderCount = 0

            If noFieldValue OrElse String.IsNullOrWhiteSpace(replaced) Then
                result.WasEmpty = True
                result.FileName = FileNameSanitizer.BuildFileName(fallbackBaseName, ".pdf", "未识别票据")
            Else
                result.FileName = FileNameSanitizer.BuildFileName(replaced, ".pdf", fallbackBaseName)
            End If

            Return result

        Catch ex As Exception
            AppLogger.Warn("命名模板渲染异常，将回退默认规则：" & ex.Message)
            result.FileName = Nothing
            Return result
        End Try
    End Function

End Module
