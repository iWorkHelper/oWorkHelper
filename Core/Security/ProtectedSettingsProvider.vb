''' <summary>
''' 敏感设置读写门面：对 My.Settings 中的 Secret Key 做透明加解密与明文迁移。
''' 只加密 Secret Key 这一敏感项；ApiKey/URL 等普通配置不加密。
''' 解密失败不抛出、不阻止插件加载；日志不输出明文。
''' </summary>
Public Module ProtectedSettingsProvider

    ''' <summary>
    ''' 读取 Secret Key 明文（供 SettingsForm 显示、OcrConfigProvider 使用）。
    ''' 解密失败返回空串，并通过 outDecryptFailed 告知调用方以便提示。
    ''' </summary>
    Public Function GetSecretKey(ByRef outDecryptFailed As Boolean) As String
        outDecryptFailed = False
        Try
            Dim stored As String = My.Settings.BaiduSecretKey
            If String.IsNullOrEmpty(stored) Then
                Return String.Empty
            End If
            Dim ok As Boolean
            Dim plain As String = SecretProtector.TryUnprotect(stored, ok)
            If Not ok Then
                outDecryptFailed = True
                Return String.Empty
            End If
            Return plain
        Catch ex As Exception
            AppLogger.Error("读取 Secret Key 异常（不影响加载）。", ex)
            outDecryptFailed = True
            Return String.Empty
        End Try
    End Function

    ''' <summary>简化重载。</summary>
    Public Function GetSecretKey() As String
        Dim failed As Boolean
        Return GetSecretKey(failed)
    End Function

    ''' <summary>
    ''' 加密并写入 Secret Key（内存），调用方负责 My.Settings.Save()。
    ''' 返回是否成功。**加密失败时不会写入**，以免把明文落到 user.config 却在日志中谎报已加密。
    ''' 注意：返回 True 仅代表内存中的值已是密文，调用方保存后仍建议调用 VerifyPersisted()。
    ''' </summary>
    Public Function SetSecretKey(plainText As String) As Boolean
        Try
            Dim cleaned As String = If(plainText, String.Empty).Trim()
            Dim ok As Boolean
            Dim enc As String = SecretProtector.TryProtect(cleaned, ok)
            If Not ok Then
                AppLogger.Error("保存 Secret Key 失败：DPAPI 加密不可用，已拒绝写入（不落明文）。")
                Return False
            End If
            My.Settings.BaiduSecretKey = enc
            Return True
        Catch ex As Exception
            AppLogger.Error("保存 Secret Key 异常（已拒绝写入）。", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 校验已持久化的 Secret Key 确实是密文。调用方在 My.Settings.Save() 之后使用，
    ''' 只有返回 True 才可记录"已加密保存"之类的成功日志。
    ''' </summary>
    Public Function VerifyPersisted() As Boolean
        Try
            Dim stored As String = My.Settings.BaiduSecretKey
            If String.IsNullOrEmpty(stored) Then Return True ' 空值无需加密
            If SecretProtector.IsProtected(stored) Then Return True
            AppLogger.Error("Secret Key 持久化校验失败：存储值不是 DPAPI 密文。")
            Return False
        Catch ex As Exception
            AppLogger.Error("Secret Key 持久化校验异常。", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 迁移：若 Secret Key 为明文遗留值，则加密回写。返回是否**确实完成**了迁移。
    ''' 加密失败时不写入、不保存，并明确记录失败（不再出现"日志说已加密、磁盘上是明文"）。
    ''' </summary>
    Public Function MigratePlaintextIfNeeded() As Boolean
        Try
            Dim stored As String = My.Settings.BaiduSecretKey
            If String.IsNullOrEmpty(stored) OrElse SecretProtector.IsProtected(stored) Then
                Return False
            End If
            ' 明文遗留 → 加密回写（失败即中止，绝不写入明文）
            Dim ok As Boolean
            Dim enc As String = SecretProtector.TryProtect(stored, ok)
            If Not ok OrElse Not SecretProtector.IsProtected(enc) Then
                AppLogger.Error("Secret Key 明文迁移失败：DPAPI 加密不可用，已保留原值未改动。")
                Return False
            End If
            My.Settings.BaiduSecretKey = enc
            My.Settings.Save()
            If Not VerifyPersisted() Then
                AppLogger.Error("Secret Key 明文迁移未生效：保存后仍非密文。")
                Return False
            End If
            AppLogger.Info("已将明文 Secret Key 迁移为 DPAPI 加密存储。")
            Return True
        Catch ex As Exception
            AppLogger.Error("Secret Key 明文迁移失败（不影响加载）。", ex)
            Return False
        End Try
    End Function

End Module
