Attribute VB_Name = "C001_Helper"
' ============================================================================
' C001_Helper.bas  —  Word 表格工作台 (Univer) VBA 辅助模块
' ============================================================================
' 功能：  处理 JS 端触发的 XLSX 保存请求
'         - 自动创建 "C001 WORD核对目录" 文件夹
'         - 从 Word 文档 Settings 读取 XLSX base64 数据
'         - 解码并写入 .xlsx 文件
'         - 根据是否本地文件决定保存位置和文件名
'
' 安装：  1. 打开 Word，按 Alt+F11 进入 VBA 编辑器
'         2. 在左侧项目资源管理器中右键 → 导入文件 → 选择此 .bas
'         3. 在 "ThisDocument" 模块中粘贴以下代码（见文件末尾）
'         4. 运行 C001_Setup 宏创建触发器 Content Control
'
' 作者：  表格工作台 (Univer)
' ============================================================================

Option Explicit

' ===== Windows API 声明 =====
Private Declare PtrSafe Function CryptStringToBinary Lib "crypt32.dll" _
    (ByVal pszString As String, ByVal cchString As Long, _
     ByVal dwFlags As Long, ByVal pbBinary As Byte, _
     ByRef pcbBinary As Long, ByRef pdwSkip As Long, _
     ByRef pdwFlags As Long) As Long

Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const CRYPT_STRING_BASE64HEADER As Long = 0
Private Const CRYPT_STRING_ANY As Long = 1

' ===== 常量 =====
Private Const C001_TRIGGER_CONTROL As String = "C001_Trigger"
Private Const C001_DIRNAME As String = "C001 WORD核对目录"

' ============================================================================
' C001_Setup — 一次性设置：创建触发器 Content Control
' ============================================================================
Public Sub C001_Setup()
    Dim fld As ContentControl
    Dim exists As Boolean
    exists = False
    
    ' 检查是否已存在
    For Each fld In ThisDocument.ContentControls
        If fld.Title = C001_TRIGGER_CONTROL Then
            exists = True
            Exit For
        End If
    Next fld
    
    If Not exists Then
        ' 在文档末尾创建一个隐藏的 Content Control
        Dim rng As Range
        Set rng = ThisDocument.Range(ThisDocument.Content.End - 1, ThisDocument.Content.End - 1)
        Set fld = ThisDocument.ContentControls.Add(wdContentControlText, rng)
        fld.Title = C001_TRIGGER_CONTROL
        fld.Tag = "C001"
        fld.Range.Font.Hidden = True
        fld.Range.Text = "READY"
        fld.LockContents = False
        
        MsgBox "C001 触发器已设置。" & vbCrLf & _
               "现在可以使用 Word → 表 按钮导入表格，" & vbCrLf & _
               "然后点击保存按钮自动生成 XLSX。", _
               vbInformation, "C001 表格工作台"
    Else
        MsgBox "C001 触发器已存在，无需重复设置。", _
               vbInformation, "C001 表格工作台"
    End If
End Sub

' ============================================================================
' C001_SaveToFolder — 核心保存逻辑（可手动触发）
' ============================================================================
Public Sub C001_SaveToFolder()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Set doc = ThisDocument
    
    ' 1. 读取文档 Settings 中的数据
    Dim xlsxB64 As String
    Dim fileName As String
    Dim targetFolder As String
    Dim isLocal As String
    
    xlsxB64 = _GetSetting(doc, "C001_XLSX_B64")
    fileName = _GetSetting(doc, "C001_Filename")
    targetFolder = _GetSetting(doc, "C001_TargetFolder")
    isLocal = _GetSetting(doc, "C001_IsLocal")
    
    If Len(xlsxB64) = 0 Then
        ' 尝试从旧格式的 Settings 读取 JSON
        Dim jsonData As String
        jsonData = _GetSetting(doc, "univerData")
        If Len(jsonData) = 0 Then
            MsgBox "没有待保存的表格数据。" & vbCrLf & _
                   "请先在 Word 中选择表格，然后点击"Word → 表"按钮导入数据。", _
                   vbInformation, "C001 表格工作台"
            Exit Sub
        End If
        MsgBox "检测到旧格式数据（Word Settings）。" & vbCrLf & _
               "请在表格工作台中点击"保存"按钮重新生成 XLSX。", _
               vbInformation, "C001 表格工作台"
        Exit Sub
    End If
    
    ' 2. 确定目标目录
    Dim savePath As String
    If Len(targetFolder) > 0 Then
        ' 使用传入的目标目录
        If Right(targetFolder, 1) <> "\" Then targetFolder = targetFolder & "\"
        savePath = targetFolder & fileName
    Else
        ' 回退：使用文档所在目录
        If doc.Path <> "" Then
            Dim docDir As String
            docDir = doc.Path
            If Right(docDir, 1) <> "\" Then docDir = docDir & "\"
            ' 创建 C001 子目录
            Dim c001Dir As String
            c001Dir = docDir & C001_DIRNAME
            If Len(Dir(c001Dir, vbDirectory)) = 0 Then
                MkDir c001Dir
            End If
            savePath = c001Dir & "\" & fileName
        Else
            ' 文档未保存，使用临时目录
            Dim tempDir As String
            tempDir = Environ("TEMP") & "\" & C001_DIRNAME
            If Len(Dir(tempDir, vbDirectory)) = 0 Then
                MkDir tempDir
            End If
            savePath = tempDir & "\" & fileName
        End If
    End If
    
    ' 3. 创建目录（如果不存在）
    Dim saveDir As String
    saveDir = Left(savePath, InStrRev(savePath, "\"))
    If Len(Dir(saveDir, vbDirectory)) = 0 Then
        MkDir saveDir
    End If
    
    ' 4. Base64 解码并写入文件
    Dim xlsxBytes() As Byte
    xlsxBytes = _Base64Decode(xlsxB64)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open savePath For Binary Access Write As #fileNum
        Put #fileNum, , xlsxBytes
    Close #fileNum
    
    ' 5. 清除 Settings
    _ClearSettings doc
    
    ' 6. 完成提示
    Dim sizeKB As Long
    sizeKB = UBound(xlsxBytes) \ 1024
    MsgBox "XLSX 文件已保存：" & vbCrLf & vbCrLf & _
           "  文件名：" & fileName & vbCrLf & _
           "  路径：" & savePath & vbCrLf & _
           "  大小：" & sizeKB & " KB" & vbCrLf & vbCrLf & _
           "请在资源管理器中打开 " & saveDir & " 查看。", _
           vbInformation, "C001 保存成功"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "保存失败（错误 " & Err.Number & "）：" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "目标路径：" & savePath, _
           vbCritical, "C001 保存错误"
End Sub

' ============================================================================
' C001_CheckAndSave — 检查是否有待保存数据，若有则保存
'                     可在 Document_Open 事件中调用
' ============================================================================
Public Sub C001_CheckAndSave()
    If _HasPendingData(ThisDocument) Then
        C001_SaveToFolder
    End If
End Sub

' ============================================================================
' 辅助函数：读取文档 Settings
' ============================================================================
Private Function _GetSetting(ByRef doc As Document, ByVal keyName As String) As String
    On Error Resume Next
    Dim val As String
    val = ""
    If doc.Settings.Exists(keyName) Then
        val = CStr(doc.Settings(keyName))
    End If
    _GetSetting = val
    On Error GoTo 0
End Function

' ============================================================================
' 辅助函数：检查是否有待保存数据
' ============================================================================
Private Function _HasPendingData(ByRef doc As Document) As Boolean
    On Error Resume Next
    Dim xlsxB64 As String
    xlsxB64 = _GetSetting(doc, "C001_XLSX_B64")
    _HasPendingData = (Len(xlsxB64) > 0)
    On Error GoTo 0
End Function

' ============================================================================
' 辅助函数：清除 C001 相关 Settings
' ============================================================================
Private Sub _ClearSettings(ByRef doc As Document)
    On Error Resume Next
    Dim keys As Variant
    Dim i As Integer
    keys = Array("C001_XLSX_B64", "C001_Filename", "C001_TargetFolder", _
                 "C001_IsLocal", "C001_Action", "C001_SheetCount", _
                 "C001_MaxRows", "C001_MaxCols")
    For i = LBound(keys) To UBound(keys)
        If doc.Settings.Exists(CStr(keys(i))) Then
            doc.Settings.Remove CStr(keys(i))
        End If
    Next i
    On Error GoTo 0
End Sub

' ============================================================================
' 辅助函数：Base64 解码
' ============================================================================
Private Function _Base64Decode(ByVal b64String As String) As Byte()
    Dim result() As Byte
    Dim outLen As Long
    Dim skip As Long
    Dim flags As Long
    
    outLen = Len(b64String)
    ReDim result(0 To outLen) As Byte
    
    CryptStringToBinary StrPtr(b64String), _
        CLng(Len(b64String)), _
        CRYPT_STRING_BASE64 Or CRYPT_STRING_ANY, _
        result(0), _
        outLen, skip, flags
    
    ReDim Preserve result(0 To outLen - 1) As Byte
    _Base64Decode = result
End Function

' ============================================================================
' 辅助函数：手动测试 Base64 解码
' ============================================================================
Public Sub C001_TestDecode()
    Dim testB64 As String
    Dim result() As Byte
    Dim i As Long
    
    testB64 = "dGVzdA=="  ' "test"
    result = _Base64Decode(testB64)
    
    Dim decoded As String
    decoded = ""
    For i = 0 To UBound(result)
        decoded = decoded & Chr(result(i))
    Next i
    
    MsgBox "Base64 解码测试：" & testB64 & " → " & decoded, _
           vbInformation, "C001 测试"
End Sub
