Imports System.Runtime.InteropServices

''' <summary>
''' COM 对象释放助手（单一实现，供全项目复用）。
'''
''' 背景：Outlook 对象模型返回的 RCW 必须显式释放，否则邮件/文件夹对象会以
''' 非确定方式驻留，长会话下会拖慢甚至拖死 Outlook。此前各文件各自维护一份私有
''' ReleaseCom 实现（ExplorerFolderService / MailAttachmentReader），易漏用、易走样。
'''
''' 约定：
'''  - 只释放"本代码自己获取"的对象；
'''  - **不要**释放 Globals.ThisAddIn.Application（VSTO 宿主项，归运行时所有）；
'''  - 释放失败一律吞掉，绝不影响主流程。
''' </summary>
Public Module ComRelease

    ''' <summary>释放单个 COM 对象（幂等、容错；非 COM 对象与 Nothing 均安全跳过）。</summary>
    Public Sub Release(comObject As Object)
        Try
            If comObject IsNot Nothing AndAlso Marshal.IsComObject(comObject) Then
                Marshal.ReleaseComObject(comObject)
            End If
        Catch
            ' 释放失败不影响主流程。
        End Try
    End Sub

    ''' <summary>依次释放多个 COM 对象（用于替代重复的 Try/Finally 样板）。</summary>
    Public Sub ReleaseAll(ParamArray comObjects As Object())
        If comObjects Is Nothing Then Return
        For Each o As Object In comObjects
            Release(o)
        Next
    End Sub

End Module
