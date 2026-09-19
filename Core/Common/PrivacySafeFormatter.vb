Imports System.IO
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

''' <summary>
''' 面向日志/报告的最小脱敏工具，默认隐藏邮件主题、文件名和本机路径中的敏感细节。
''' </summary>
Public Module PrivacySafeFormatter

    ''' <summary>匹配盘符路径（C:\...）或 UNC 路径（\\server\share\...）。</summary>
    Private ReadOnly PathPattern As New Regex(
        "([A-Za-z]:\\[^""'\s;，。、）)]*)|(\\\\[^""'\s;，。、）)]+)",
        RegexOptions.Compiled)

    ''' <summary>
    ''' 清除自由文本中的本机路径痕迹。用于异常消息等"非结构化"文案：
    ''' 这类消息常内嵌完整临时路径，而临时路径里又嵌有原始附件文件名，
    ''' 直接写入报告/日志会绕过 MaskPath/MaskFileName 的脱敏口径。
    ''' </summary>
    Public Function ScrubPaths(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        Try
            Return PathPattern.Replace(text, AddressOf ScrubPathMatch)
        Catch
            Return text
        End Try
    End Function

    Private Function ScrubPathMatch(m As Match) As String
        Return MaskPath(m.Value)
    End Function

    Public Function MaskSubject(subject As String) As String
        If String.IsNullOrWhiteSpace(subject) Then
            Return "(无主题)"
        End If
        Return MaskText(subject.Trim(), 2, 2)
    End Function

    Public Function MaskFileName(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then
            Return "(空)"
        End If

        Dim ext As String = String.Empty
        Dim baseName As String = fileName.Trim()
        Try
            ext = Path.GetExtension(baseName)
            baseName = Path.GetFileNameWithoutExtension(baseName)
        Catch
        End Try

        Return MaskText(baseName, 1, 1) & ext
    End Function

    Public Function MaskPath(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then
            Return "(空)"
        End If

        Try
            Dim trimmed As String = pathValue.Trim()
            Dim fileName As String = Path.GetFileName(trimmed)
            If String.IsNullOrWhiteSpace(fileName) Then
                Return "<path>"
            End If

            If trimmed.StartsWith("\\", StringComparison.Ordinal) Then
                Return "\\<share>\...\" & MaskFileName(fileName)
            End If

            Dim root As String = Path.GetPathRoot(trimmed)
            If Not String.IsNullOrWhiteSpace(root) Then
                Return root & "...\" & MaskFileName(fileName)
            End If

            Return "...\\" & MaskFileName(fileName)
        Catch
            Return "<path>"
        End Try
    End Function

    Public Function MaskFileList(items As IEnumerable(Of String)) As String
        If items Is Nothing Then
            Return "-"
        End If

        Dim masked As New List(Of String)()
        For Each item As String In items
            masked.Add(MaskFileName(item))
        Next

        If masked.Count = 0 Then
            Return "-"
        End If
        Return String.Join(", ", masked.ToArray())
    End Function

    Private Function MaskText(value As String, keepPrefix As Integer, keepSuffix As Integer) As String
        If String.IsNullOrWhiteSpace(value) Then
            Return String.Empty
        End If

        Dim trimmed As String = value.Trim()
        If trimmed.Length <= 2 Then
            Return New String("*"c, trimmed.Length)
        End If

        Dim prefixLen As Integer = Math.Max(1, Math.Min(keepPrefix, trimmed.Length - 1))
        Dim suffixLen As Integer = Math.Max(0, Math.Min(keepSuffix, trimmed.Length - prefixLen))
        If prefixLen + suffixLen >= trimmed.Length Then
            suffixLen = Math.Max(0, trimmed.Length - prefixLen - 1)
        End If

        Dim middleLen As Integer = Math.Max(1, trimmed.Length - prefixLen - suffixLen)
        Dim prefix As String = trimmed.Substring(0, prefixLen)
        Dim suffix As String = If(suffixLen > 0, trimmed.Substring(trimmed.Length - suffixLen), String.Empty)
        Return prefix & New String("*"c, middleLen) & suffix
    End Function

End Module
