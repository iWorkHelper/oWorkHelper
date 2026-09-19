Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' 敏感字符串加解密（Windows DPAPI，CurrentUser 作用域）。
''' 安全边界：DPAPI 仅保护"当前 Windows 用户"下的本地数据，换用户/换机器无法解密；
''' 不能替代访问控制/权限管理。仅用于加密 Secret Key 等敏感项，普通配置不加密。
''' 加密结果带 "DPAPI:" 前缀以便识别与迁移。日志中不得输出明文。
''' </summary>
Public Module SecretProtector

    Public Const Prefix As String = "DPAPI:"

    ''' <summary>是否为已加密（带前缀）的值。</summary>
    Public Function IsProtected(value As String) As Boolean
        Return Not String.IsNullOrEmpty(value) AndAlso value.StartsWith(Prefix, StringComparison.Ordinal)
    End Function

    ''' <summary>
    ''' 加密明文，返回带前缀的 Base64。空值原样返回（不加密空串）。
    ''' **失败返回 Nothing，绝不回退为明文**：调用方应先检查成功标志（TryProtect），
    ''' 失败时拒绝持久化。原实现失败时返回原文，会在日志谎报"已加密"的同时把明文写入磁盘。
    ''' </summary>
    Public Function Protect(plainText As String) As String
        Dim ok As Boolean
        Return TryProtect(plainText, ok)
    End Function

    ''' <summary>
    ''' 加密明文并返回是否成功。成功得到带 "DPAPI:" 前缀的 Base64；
    ''' 空值视为成功并原样返回；已是密文视为成功并原样返回；失败返回 Nothing 且 success=False。
    ''' </summary>
    Public Function TryProtect(plainText As String, ByRef success As Boolean) As String
        success = False
        If String.IsNullOrEmpty(plainText) Then
            success = True
            Return plainText ' 空串无需加密
        End If
        If IsProtected(plainText) Then
            success = True
            Return plainText ' 已是密文，避免重复加密
        End If
        Try
            Dim bytes As Byte() = Encoding.UTF8.GetBytes(plainText)
            Dim encrypted As Byte() = ProtectedData.Protect(bytes, Nothing, DataProtectionScope.CurrentUser)
            success = True
            Return Prefix & Convert.ToBase64String(encrypted)
        Catch ex As Exception
            AppLogger.Error("Secret 加密失败（已放弃写入，绝不回退为明文）。", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 解密。若非前缀（明文遗留），原样返回。解密失败返回空并标记失败。
    ''' </summary>
    Public Function TryUnprotect(storedValue As String, ByRef success As Boolean) As String
        success = True
        If String.IsNullOrEmpty(storedValue) Then
            Return storedValue
        End If
        If Not IsProtected(storedValue) Then
            ' 明文遗留值：原样返回（供迁移）。
            Return storedValue
        End If
        Try
            Dim base64 As String = storedValue.Substring(Prefix.Length)
            Dim encrypted As Byte() = Convert.FromBase64String(base64)
            Dim bytes As Byte() = ProtectedData.Unprotect(encrypted, Nothing, DataProtectionScope.CurrentUser)
            Return Encoding.UTF8.GetString(bytes)
        Catch ex As Exception
            success = False
            AppLogger.Error("Secret 解密失败（可能来自其它用户/机器的配置）。", ex)
            Return String.Empty
        End Try
    End Function

    ''' <summary>解密（忽略成功标志的简化重载）。</summary>
    Public Function Unprotect(storedValue As String) As String
        Dim ok As Boolean
        Return TryUnprotect(storedValue, ok)
    End Function

End Module
