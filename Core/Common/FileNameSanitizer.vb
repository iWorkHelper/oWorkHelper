Imports System.IO
Imports System.Text

''' <summary>
''' 文件名清理工具：去除 Windows 非法字符、处理保留名、限制长度。
''' 用于把从发票中提取的字段安全地拼接为文件名。
''' </summary>
Public Module FileNameSanitizer

    ''' <summary>Windows 文件名最大安全长度（保守值，避免路径整体超限）。</summary>
    Private Const MaxBaseNameLength As Integer = 120

    ''' <summary>Windows 保留设备名，不能作为文件名主体。</summary>
    Private ReadOnly ReservedNames As String() = New String() {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"}

    ''' <summary>
    ''' 清理文件名主体（不含扩展名）。非法字符替换为下划线；
    ''' 空白折叠；保留名加前缀；超长截断。
    ''' </summary>
    ''' <param name="rawName">原始名称片段。</param>
    ''' <param name="fallback">清理后为空时使用的回退名。</param>
    Public Function SanitizeBaseName(rawName As String, Optional fallback As String = "未命名") As String
        If String.IsNullOrWhiteSpace(rawName) Then
            Return fallback
        End If

        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim sb As New StringBuilder(rawName.Length)
        For Each c As Char In rawName
            If Array.IndexOf(invalid, c) >= 0 Then
                ' Path.GetInvalidFileNameChars() 已包含 0x00-0x1F（含 Tab/LF/CR），
                ' 因此控制字符在此统一处理：控制字符替换为空格（保留词边界），
                ' 其余非法字符替换为下划线。原先单独的 Tab/Lf/Cr 分支永远不可达（死代码）。
                sb.Append(If(Char.IsControl(c), " "c, "_"c))
            Else
                sb.Append(c)
            End If
        Next

        Dim cleaned As String = CollapseWhitespace(sb.ToString())

        ' 去除首尾空白与点（Windows 不允许结尾为点/空格）。
        cleaned = SafeTrimDotsAndSpaces(cleaned)

        If cleaned.Length = 0 Then
            Return fallback
        End If

        ' 保留名处理：Windows 判定的是"第一个点之前"的片段，
        ' 因此 NUL.pdf / COM1.invoice 同样是保留设备名，不能只比较完全相等。
        If IsReservedDeviceName(cleaned) Then
            cleaned = "_" & cleaned
        End If

        ' 超长截断（截断后仍可能留下结尾点/空格，需再次清理）。
        If cleaned.Length > MaxBaseNameLength Then
            cleaned = SafeTrimDotsAndSpaces(cleaned.Substring(0, MaxBaseNameLength))
        End If

        If cleaned.Length = 0 Then
            Return fallback
        End If

        Return cleaned
    End Function

    ''' <summary>折叠连续空白为单个空格（单遍完成，避免原先的 O(n²) Replace 循环）。</summary>
    Private Function CollapseWhitespace(value As String) As String
        If String.IsNullOrEmpty(value) Then Return String.Empty
        Dim sb As New StringBuilder(value.Length)
        Dim lastWasSpace As Boolean = False
        For Each c As Char In value
            Dim isSpace As Boolean = Char.IsWhiteSpace(c)
            If isSpace Then
                If Not lastWasSpace Then sb.Append(" "c)
            Else
                sb.Append(c)
            End If
            lastWasSpace = isSpace
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 反复去除首尾的空白与点，直到稳定。
    ''' 原先的 Trim().Trim("."c).Trim() 单次链式处理会留下结尾点
    ''' （"abc. ." → "abc."），而 Windows 会静默丢弃结尾点，
    ''' 造成期望文件名与实际文件名不一致、甚至两个不同名称互相冲突。
    ''' </summary>
    Private Function SafeTrimDotsAndSpaces(value As String) As String
        Dim s As String = If(value, String.Empty)
        Dim changed As Boolean = True
        While changed AndAlso s.Length > 0
            changed = False
            If s.Length > 0 AndAlso (Char.IsWhiteSpace(s(0)) OrElse s(0) = "."c) Then
                s = s.Substring(1)
                changed = True
            End If
            If s.Length > 0 AndAlso (Char.IsWhiteSpace(s(s.Length - 1)) OrElse s(s.Length - 1) = "."c) Then
                s = s.Substring(0, s.Length - 1)
                changed = True
            End If
        End While
        Return s
    End Function

    ''' <summary>是否为 Windows 保留设备名（按"第一个点之前"的片段判定）。</summary>
    Private Function IsReservedDeviceName(value As String) As Boolean
        Dim stem As String = If(value, String.Empty)
        Dim dot As Integer = stem.IndexOf("."c)
        If dot >= 0 Then stem = stem.Substring(0, dot)
        stem = stem.TrimEnd(" "c).ToUpperInvariant()
        For Each reserved As String In ReservedNames
            If String.Equals(stem, reserved, StringComparison.Ordinal) Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' 生成完整文件名（自动补全扩展名，默认 .pdf）。
    ''' </summary>
    Public Function BuildFileName(baseName As String, Optional extension As String = ".pdf", Optional fallback As String = "未命名") As String
        Dim safeBase As String = SanitizeBaseName(baseName, fallback)
        Dim ext As String = If(extension, String.Empty)
        If ext.Length > 0 AndAlso Not ext.StartsWith(".", StringComparison.Ordinal) Then
            ext = "." & ext
        End If
        Return safeBase & ext
    End Function

End Module
