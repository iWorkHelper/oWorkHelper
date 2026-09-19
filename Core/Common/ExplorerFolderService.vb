Imports System.IO
Imports System.Runtime.InteropServices

''' <summary>
''' 归档目录“打开 / 激活”服务。
''' 仅在批量归档流程正常结束、且本批确有文件成功归档后由调用方执行一次：
'''  - 若目标目录已在“文件资源管理器”中打开 → 恢复（若最小化）并激活现有窗口，不新开窗口；
'''  - 若尚未打开 → 用系统资源管理器打开该目录。
'''
''' 关键实现要点（修复“已打开仍新开窗口”的 Bug）：
'''  - 通过 Shell.Application.Windows() 枚举 Shell 窗口；每个窗口的真实目录优先取
'''    window.Document.Folder.Self.Path（最可靠），不可用时再回退 LocationURL → 本地路径；
'''    不用窗口标题 / LocationName / 进程名判断。
'''  - 目标路径与窗口路径统一经 NormalizeFolderPath 规范化后忽略大小写比较
'''    （URI 解码、file:/// 转本地、GetFullPath、/→\、去尾分隔符但保留盘符根、支持中文/空格/UNC）。
'''  - 控制流严格分离“查找命中”与“激活成功”：只要查找到匹配窗口即 Return，绝不再走新开逻辑；
'''    激活 API 失败只写日志。
'''  - 窗口句柄取各 Shell 窗口对象自身的 HWND（不是进程主窗口句柄）。
'''  - 单个窗口读取异常不中断枚举；COM 对象在 Finally 中释放，避免 Outlook 长期持有引用。
''' </summary>
Public Module ExplorerFolderService

    ' —— Win32 API ——
    <DllImport("user32.dll")>
    Private Function IsIconic(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function ShowWindowAsync(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function BringWindowToTop(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function SetFocus(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <DllImport("kernel32.dll")>
    Private Function GetCurrentThreadId() As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Function AttachThreadInput(idAttach As UInteger, idAttachTo As UInteger, fAttach As Boolean) As Boolean
    End Function

    Private Const SW_RESTORE As Integer = 9
    Private Const SW_SHOW As Integer = 5

    ''' <summary>
    ''' 打开或激活指定目录：已打开则激活现有窗口，未打开则新开资源管理器窗口。
    ''' 目录为空 / 非法 / 不存在时仅记录并跳过（不自动创建目录）。
    ''' </summary>
    Public Sub OpenOrActivateFolder(folderPath As String)
        Try
            If String.IsNullOrWhiteSpace(folderPath) Then
                AppLogger.Warn("打开归档目录：目录为空，已跳过。")
                Return
            End If

            ' 用于存在性判断与新开的“可用本地路径”。
            Dim localPath As String
            Try
                localPath = Path.GetFullPath(folderPath.Trim())
            Catch ex As Exception
                AppLogger.Warn("打开归档目录：路径非法，已跳过。" & ex.Message)
                Return
            End Try

            Dim targetNorm As String = NormalizeFolderPath(folderPath)
            AppLogger.Info("打开/激活归档目录：原始=" & folderPath & "；规范化=" & targetNorm)

            ' 目录不存在 → 告警并跳过（不自动创建；创建职责属于归档流程本身）。
            If Not Directory.Exists(localPath) Then
                AppLogger.Warn("打开归档目录：目录不存在，已跳过（不自动创建）：" & localPath)
                Return
            End If

            ' 1) 查找是否已有资源管理器窗口打开该目录。
            Dim hwnd As IntPtr = IntPtr.Zero
            If TryFindExplorerWindow(targetNorm, hwnd) Then
                AppLogger.Info("已匹配到打开该目录的资源管理器窗口（句柄=" & hwnd.ToInt64().ToString() & "），执行激活（不新开）。")
                ActivateExplorerWindow(hwnd)
                Return   ' 只要查找到匹配窗口，无论激活成功与否都不再新开。
            End If

            ' 2) 未找到 → 新开窗口。
            AppLogger.Info("未找到打开该目录的资源管理器窗口，执行新开。")
            OpenNewExplorerWindow(localPath)
        Catch ex As Exception
            ' 兜底：绝不因打开/激活目录影响归档主流程。
            AppLogger.Warn("打开/激活归档目录发生异常（已忽略，不影响归档）：" & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 枚举 Shell 窗口，查找“当前显示目录”与目标规范化路径一致的资源管理器窗口。
    ''' 命中则通过 ByRef 返回其窗口句柄并返回 True；否则返回 False。
    ''' </summary>
    Private Function TryFindExplorerWindow(targetNorm As String, ByRef explorerHwnd As IntPtr) As Boolean
        explorerHwnd = IntPtr.Zero
        Dim shellApp As Object = Nothing
        Dim windows As Object = Nothing
        ' O-27/O-33：窗口对象先复制到托管列表，枚举期间不释放元素；
        ' 枚举结束后（含命中提前返回）再统一释放各窗口 RCW。
        Dim windowItems As New List(Of Object)()
        Try
            Dim shellType As Type = Type.GetTypeFromProgID("Shell.Application")
            If shellType Is Nothing Then
                AppLogger.Warn("Shell.Application ProgID 未注册，无法枚举窗口，将改为新开。")
                Return False
            End If
            shellApp = Activator.CreateInstance(shellType)
            windows = shellApp.Windows()
            If windows Is Nothing Then Return False

            AppLogger.Debug("开始枚举 Shell 窗口（快照后遍历）；目标规范化路径=" & targetNorm)

            windowItems = SnapshotComCollection(windows)

            Dim idx As Integer = -1
            ' 注意：ShellWindows 集合必须整体枚举后才能安全释放元素；对其做整数索引（Item(i)）在
            ' 后期绑定下会取不到窗口对象，曾导致“已打开仍新开窗口”。
            For Each w As Object In windowItems
                idx += 1
                Try
                    If w Is Nothing Then Continue For

                    ' 过滤非文件资源管理器窗口（如 Internet Explorer）。
                    If Not IsFileExplorerWindow(w) Then Continue For

                    Dim winPath As String = Nothing
                    Dim gotPath As Boolean = TryGetExplorerFolderPath(w, winPath)
                    Dim locUrl As String = SafeGetString(Function() CStr(w.LocationURL))

                    If Not gotPath Then
                        AppLogger.Debug(String.Format("  窗口[{0}] LocationURL={1}；无法取得文件夹路径，跳过。", idx, Dash(locUrl)))
                        Continue For
                    End If

                    Dim winNorm As String = NormalizeFolderPath(winPath)
                    Dim isMatch As Boolean = (winNorm.Length > 0) AndAlso
                        String.Equals(winNorm, targetNorm, StringComparison.OrdinalIgnoreCase)

                    AppLogger.Debug(String.Format("  窗口[{0}] LocationURL={1}；Folder.Self.Path={2}；规范化={3}；匹配={4}",
                                                  idx, Dash(locUrl), Dash(winPath), Dash(winNorm), isMatch))

                    If isMatch Then
                        Dim hwnd As IntPtr = TryGetHwnd(w)
                        AppLogger.Debug("  命中窗口，HWND=" & hwnd.ToInt64().ToString())
                        If hwnd <> IntPtr.Zero Then
                            explorerHwnd = hwnd
                            Return True
                        Else
                            AppLogger.Debug("  命中窗口但 HWND 无效，继续枚举其它窗口。")
                        End If
                    End If
                Catch exItem As Exception
                    AppLogger.Debug("  枚举单个窗口异常（跳过该窗口）：" & exItem.Message)
                End Try
            Next
            Return False
        Catch ex As Exception
            AppLogger.Warn("枚举资源管理器窗口失败（将改为新开）：" & ex.Message)
            Return False
        Finally
            ReleaseComAll(windowItems)
            ReleaseCom(windows)
            ReleaseCom(shellApp)
        End Try
    End Function

    ''' <summary>
    ''' 把 COM 集合（IEnumVARIANT）安全复制为托管列表：
    '''  - 枚举期间**不释放**元素 RCW，避免边枚举边释放导致集合失效；
    '''  - 显式持有并释放枚举器 RCW（O-33），失败时回退到 For Each。
    ''' </summary>
    Private Function SnapshotComCollection(collection As Object) As List(Of Object)
        Dim items As New List(Of Object)()
        If collection Is Nothing Then Return items

        Dim enumerator As System.Collections.IEnumerator = Nothing
        Try
            ' .NET 为实现了 IEnumVARIANT 的 COM 集合提供 IEnumerable 适配器。
            enumerator = DirectCast(collection, System.Collections.IEnumerable).GetEnumerator()
            While enumerator.MoveNext()
                items.Add(enumerator.Current)
            End While
            Return items
        Catch ex As Exception
            ' 回退：常规 For Each（同样保证枚举期间不释放元素）。
            AppLogger.Debug("COM 集合显式枚举失败，回退 For Each：" & ex.Message)
            items.Clear()
            For Each w As Object In collection
                items.Add(w)
            Next
            Return items
        Finally
            ReleaseCom(enumerator)
        End Try
    End Function

    ''' <summary>释放列表中的全部 COM 对象（仅在枚举结束后调用）。</summary>
    Private Sub ReleaseComAll(items As List(Of Object))
        If items Is Nothing Then Return
        For Each o As Object In items
            ReleaseCom(o)
        Next
    End Sub

    ''' <summary>
    ''' 判断是否为文件资源管理器窗口（排除 IE）。
    ''' FullName 不可用时不据此排除，交由能否取得真实文件夹路径判断，避免因 FullName 读取异常
    ''' 而漏掉正常的资源管理器窗口（Win11 曾出现此类漏配导致每次都新开窗口）。
    ''' </summary>
    Private Function IsFileExplorerWindow(w As Object) As Boolean
        Dim fullName As String = SafeGetString(Function() CStr(w.FullName))
        If String.IsNullOrEmpty(fullName) Then Return True
        Dim exe As String
        Try
            exe = Path.GetFileName(fullName)
        Catch
            exe = fullName
        End Try
        ' 明确排除 IE；explorer.exe 及其它 shell 宿主一律放行（最终以能否取得文件夹路径为准）。
        If String.Equals(exe, "iexplore.exe", StringComparison.OrdinalIgnoreCase) Then Return False
        Return True
    End Function

    ''' <summary>
    ''' 取窗口当前显示目录的真实本地路径。
    ''' 优先 window.Document.Folder.Self.Path；不可用时回退 LocationURL → 本地路径。
    ''' </summary>
    Private Function TryGetExplorerFolderPath(w As Object, ByRef folderPath As String) As Boolean
        folderPath = Nothing

        ' —— 首选：Document.Folder.Self.Path ——
        Dim doc As Object = Nothing
        Dim folder As Object = Nothing
        Dim self As Object = Nothing
        Try
            doc = w.Document
            If doc IsNot Nothing Then
                folder = doc.Folder
                If folder IsNot Nothing Then
                    self = folder.Self
                    If self IsNot Nothing Then
                        Dim p As String = SafeGetString(Function() CStr(self.Path))
                        If Not String.IsNullOrWhiteSpace(p) Then folderPath = p
                    End If
                End If
            End If
        Catch exDoc As Exception
            AppLogger.Debug("    读取 Document.Folder.Self.Path 失败：" & exDoc.Message)
        Finally
            ReleaseCom(self)
            ReleaseCom(folder)
            ReleaseCom(doc)
        End Try

        If Not String.IsNullOrWhiteSpace(folderPath) Then Return True

        ' —— 回退：LocationURL → 本地路径（含 URI 解码）——
        Try
            Dim locUrl As String = SafeGetString(Function() CStr(w.LocationURL))
            If Not String.IsNullOrEmpty(locUrl) Then
                Dim u As New Uri(locUrl)
                If u.IsFile Then folderPath = u.LocalPath
            End If
        Catch exUrl As Exception
            AppLogger.Debug("    LocationURL 转本地路径失败：" & exUrl.Message)
        End Try

        Return Not String.IsNullOrWhiteSpace(folderPath)
    End Function

    ''' <summary>
    ''' 统一规范化目录路径：URI 解码 / file:/// 转本地 / GetFullPath / (/→\) / 去尾分隔符（保留盘符根）。
    ''' 支持中文、空格、%20 等 URI 编码、UNC 路径、盘符大小写差异（比较时用 OrdinalIgnoreCase）。
    ''' </summary>
    Private Function NormalizeFolderPath(rawPath As String) As String
        If String.IsNullOrWhiteSpace(rawPath) Then Return String.Empty
        Dim p As String = rawPath.Trim()
        Try
            ' 1) file:/// URI → 本地路径（并完成 URI 解码）。
            If p.StartsWith("file:", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim u As New Uri(p)
                    If u.IsFile Then
                        p = u.LocalPath
                    Else
                        p = Uri.UnescapeDataString(p)
                    End If
                Catch
                    p = Uri.UnescapeDataString(p)
                End Try
            ElseIf p.IndexOf("%"c) >= 0 Then
                ' 含 %20 等编码但非 file: 前缀 → 解码。
                Try
                    p = Uri.UnescapeDataString(p)
                Catch
                End Try
            End If

            ' 2) 分隔符统一 / → \。
            p = p.Replace("/"c, "\"c)

            ' 3) GetFullPath 规范化（解析 .、..、相对段等）。
            Try
                p = Path.GetFullPath(p)
            Catch
            End Try

            ' 4) 去尾分隔符，但保留盘符根（C:\）与 UNC 根。
            p = TrimTrailingSeparators(p)
            Return p
        Catch
            Return If(p, String.Empty)
        End Try
    End Function

    ''' <summary>去除末尾 \ 或 /，但不破坏盘符根（C:\）。</summary>
    Private Function TrimTrailingSeparators(p As String) As String
        If String.IsNullOrEmpty(p) Then Return p
        ' 盘符根 "C:\" 原样保留。
        If p.Length = 3 AndAlso Char.IsLetter(p(0)) AndAlso p(1) = ":"c AndAlso p(2) = "\"c Then
            Return p
        End If
        Dim trimmed As String = p.TrimEnd("\"c, "/"c)
        If trimmed.Length = 0 Then Return p
        ' 若截成 "C:" → 补回根 "C:\"。
        If trimmed.Length = 2 AndAlso Char.IsLetter(trimmed(0)) AndAlso trimmed(1) = ":"c Then
            Return trimmed & "\"
        End If
        Return trimmed
    End Function

    ''' <summary>
    ''' 恢复（若最小化）并将窗口激活到前台。
    ''' SetForegroundWindow 因 Windows 前台限制返回失败时，用 AttachThreadInput 兜底再尝试；
    ''' 仍失败也只记日志，绝不因此新开窗口。
    ''' </summary>
    Private Sub ActivateExplorerWindow(hwnd As IntPtr)
        Try
            If hwnd = IntPtr.Zero Then Return

            If IsIconic(hwnd) Then
                ShowWindow(hwnd, SW_RESTORE)
            Else
                ShowWindow(hwnd, SW_SHOW)
            End If
            BringWindowToTop(hwnd)

            If SetForegroundWindow(hwnd) Then
                Return
            End If

            ' —— AttachThreadInput 兜底 ——
            AppLogger.Debug("SetForegroundWindow 直接调用失败，改用 AttachThreadInput 兜底激活。")
            Dim foreground As IntPtr = GetForegroundWindow()
            Dim curThread As UInteger = GetCurrentThreadId()
            Dim dummyPid As UInteger = 0UI
            Dim foreThread As UInteger = 0UI
            Dim targetThread As UInteger = 0UI
            If foreground <> IntPtr.Zero Then foreThread = GetWindowThreadProcessId(foreground, dummyPid)
            targetThread = GetWindowThreadProcessId(hwnd, dummyPid)

            Dim attachedFore As Boolean = False
            Dim attachedTarget As Boolean = False
            Try
                If foreThread <> 0UI AndAlso foreThread <> curThread Then attachedFore = AttachThreadInput(curThread, foreThread, True)
                If targetThread <> 0UI AndAlso targetThread <> curThread AndAlso targetThread <> foreThread Then attachedTarget = AttachThreadInput(curThread, targetThread, True)

                ShowWindow(hwnd, SW_SHOW)
                BringWindowToTop(hwnd)
                Dim ok As Boolean = SetForegroundWindow(hwnd)
                SetFocus(hwnd)
                If Not ok Then
                    ShowWindowAsync(hwnd, SW_SHOW)
                    AppLogger.Warn("SetForegroundWindow 仍返回失败（Windows 前台限制），已尽力激活；仍不新开窗口。")
                End If
            Finally
                If attachedTarget Then AttachThreadInput(curThread, targetThread, False)
                If attachedFore Then AttachThreadInput(curThread, foreThread, False)
            End Try
        Catch ex As Exception
            ' 激活失败只记日志，不影响归档、也不新开窗口。
            AppLogger.Warn("激活资源管理器窗口异常（已忽略，不新开）：" & ex.Message)
        End Try
    End Sub

    ''' <summary>新开一个资源管理器窗口打开指定目录。</summary>
    Private Sub OpenNewExplorerWindow(folderPath As String)
        Try
            ' O-26：Process.Start 返回的 Process 实现 IDisposable，必须释放（否则句柄/资源随 GC 才回收）。
            Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start("explorer.exe", """" & folderPath & """")
            If p IsNot Nothing Then p.Dispose()
            AppLogger.Info("已新开资源管理器窗口打开归档目录：" & folderPath)
        Catch ex As Exception
            AppLogger.Warn("新开资源管理器窗口失败（不影响归档）：" & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 诊断用：枚举当前 Shell 窗口并报告与目标路径的匹配情况（供 OfflineTester 在真实环境验证匹配逻辑）。
    ''' 仅读取、不激活、不新开窗口；生产代码不调用。
    ''' </summary>
    Public Function BuildScanReport(targetPath As String) As String
        Dim sb As New System.Text.StringBuilder()
        Dim targetNorm As String = NormalizeFolderPath(targetPath)
        sb.AppendLine("目标原始=" & If(targetPath, "-"))
        sb.AppendLine("目标规范化=" & targetNorm)
        Dim shellApp As Object = Nothing
        Dim windows As Object = Nothing
        Dim windowItems As New List(Of Object)()
        Dim matchedHwnd As IntPtr = IntPtr.Zero
        Try
            Dim shellType As Type = Type.GetTypeFromProgID("Shell.Application")
            If shellType Is Nothing Then
                sb.AppendLine("Shell.Application 不可用")
                Return sb.ToString()
            End If
            shellApp = Activator.CreateInstance(shellType)
            windows = shellApp.Windows()
            Dim count As Integer = 0
            Try
                count = CInt(windows.Count)
            Catch
            End Try
            sb.AppendLine("Shell 窗口数=" & count)
            ' 先快照再遍历（枚举期间不释放元素；枚举器 RCW 亦被释放）。
            windowItems = SnapshotComCollection(windows)
            Dim idx As Integer = -1
            For Each w As Object In windowItems
                idx += 1
                Try
                    If w Is Nothing Then Continue For
                    Dim isExp As Boolean = IsFileExplorerWindow(w)
                    Dim locUrl As String = SafeGetString(Function() CStr(w.LocationURL))
                    Dim winPath As String = Nothing
                    Dim got As Boolean = TryGetExplorerFolderPath(w, winPath)
                    Dim winNorm As String = If(got, NormalizeFolderPath(winPath), "")
                    Dim isMatch As Boolean = got AndAlso winNorm.Length > 0 AndAlso
                        String.Equals(winNorm, targetNorm, StringComparison.OrdinalIgnoreCase)
                    sb.AppendLine(String.Format("[{0}] 资源管理器={1}; LocationURL={2}; Folder.Self.Path={3}; 规范化={4}; 匹配={5}",
                                                idx, isExp, Dash(locUrl), Dash(winPath), Dash(winNorm), isMatch))
                    If isMatch AndAlso matchedHwnd = IntPtr.Zero Then matchedHwnd = TryGetHwnd(w)
                Catch ex As Exception
                    sb.AppendLine(String.Format("[{0}] 读取异常：{1}", idx, ex.Message))
                End Try
            Next
            If matchedHwnd <> IntPtr.Zero Then
                sb.AppendLine("决策=激活已有窗口，HWND=" & matchedHwnd.ToInt64().ToString())
            Else
                sb.AppendLine("决策=新开窗口")
            End If
        Catch ex As Exception
            sb.AppendLine("枚举失败：" & ex.Message)
        Finally
            ReleaseComAll(windowItems)
            ReleaseCom(windows)
            ReleaseCom(shellApp)
        End Try
        Return sb.ToString()
    End Function

    ''' <summary>诊断用：返回路径规范化结果（供 OfflineTester 校验中文/空格/URI/大小写/尾分隔符处理）。</summary>
    Public Function NormalizeForDiagnostics(path As String) As String
        Return NormalizeFolderPath(path)
    End Function

    ''' <summary>取 Shell 窗口对象自身的 HWND（对应该目录窗口，而非进程主窗口）。</summary>
    Private Function TryGetHwnd(w As Object) As IntPtr
        Try
            Return New IntPtr(Convert.ToInt64(w.HWND))
        Catch
            Return IntPtr.Zero
        End Try
    End Function

    Private Function SafeGetString(getter As Func(Of String)) As String
        Try
            Return getter()
        Catch
            Return Nothing
        End Try
    End Function

    Private Function Dash(s As String) As String
        Return If(String.IsNullOrEmpty(s), "-", s)
    End Function

    ''' <summary>安全释放 COM 对象（统一委托到 ComRelease，避免多处实现走样）。</summary>
    Private Sub ReleaseCom(o As Object)
        ComRelease.Release(o)
    End Sub

End Module
