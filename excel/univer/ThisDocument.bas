' ============================================================================
' ThisDocument.bas  —  需要粘贴到 Word 文档 "ThisDocument" 模块的代码
' ============================================================================
' 步骤：
'   1. 打开 Word，按 Alt+F11 进入 VBA 编辑器
'   2. 在左侧项目资源管理器中，找到当前文档的 "ThisDocument"
'   3. 双击 "ThisDocument" 打开代码窗口
'   4. 将下面的代码完整粘贴进去
'   5. 保存文档并关闭 VBA 编辑器
'   6. 运行 C001_Setup 宏（开发工具 → 宏 → C001_Setup）创建触发器
' ============================================================================

Option Explicit

' 在文档打开时检查是否有未完成的保存请求
Private Sub Document_Open()
    C001_CheckAndSave
End Sub

' 在新建文档时检查
Private Sub Document_New()
    C001_CheckAndSave
End Sub

' Content Control 退出事件 — 用于触发 VBA 保存
' 当 JS 端写入 C001_Trigger 控件时，此事件会被触发
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, ByVal Cancel As Boolean)
    On Error Resume Next
    
    If ContentControl.Title = "C001_Trigger" Then
        ' 触发 C001 保存流程
        C001_SaveToFolder
    End If
    
    On Error GoTo 0
End Sub
