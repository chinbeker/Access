Attribute VB_Name = "Core"
'@Lang VBA

Option Compare Database
Option Explicit

' 窗体是否已打开
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    If VBA.Len(FormName) > 0 Then
        IsFormOpen = (Application.SysCmd(acSysCmdGetObjectState, acForm, FormName) <> 0)
    Else
        IsFormOpen = False
    End If
End Function

'打开指定窗体（统一验证授权）
Public Sub OpenForm(ByVal FormName As String, Optional ByVal Condition As String, Optional ByVal OpenArgs As String)
    On Error GoTo ErrorHandler
    If VBA.Len(FormName) = 0 Then Exit Sub
    If Not Auth.Authorized Then
        Call Auth.RedirectToLogin(FormName)
    Else
        DoCmd.OpenForm FormName, , , Condition, , , OpenArgs
    End If
    Exit Sub

ErrorHandler:
    Call Message.Warning("打开失败:                                                  " & vbCrLf & vbCrLf & "窗体 " & FormName & " 未找到")
    Exit Sub
End Sub

' 关闭指定窗体
Public Sub CloseForm(ByVal FormName As String)
    On Error GoTo ErrorHandler
    If IsFormOpen(FormName) Then DoCmd.Close acForm, FormName, acSaveNo
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

'获取指定窗体（已打开）
Public Function GetForm(ByVal FormName As String) As Form
    If Core.IsFormOpen(FormName) Then
        Set GetForm = Forms.Item(FormName)
    Else
        Set GetForm = Nothing
    End If
End Function

' 激活指定窗体
Public Sub FormSetFocus(ByVal FormName As String)
    On Error GoTo ErrorHandler
    If IsFormOpen(FormName) Then Forms.Item(FormName).SetFocus
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub


' 跳转到指定记录
Public Function GoToRecord(ByVal FormName As String, ByVal Criteria As String) As Boolean
    On Error GoTo ErrorHandler
    If Not IsFormOpen(FormName) Then Exit Function

    Dim Target As Form
    Dim rs As DAO.Recordset

    Set Target = Forms.Item(FormName)
    If Target Is Nothing Then Exit Function

    Set rs = Target.RecordsetClone
    If rs Is Nothing Then Exit Function

    If Not rs.EOF And Not rs.BOF Then
        If Not StringBase.IsNullOrWhiteSpace(Criteria) Then
            rs.FindFirst Criteria
            If Not rs.NoMatch Then
                Target.Bookmark = rs.Bookmark
                GoToRecord = True
            End If
        End If
    End If

    Target.SetFocus
    rs.Close
    Set Target = Nothing
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    If Not Target Is Nothing Then Set Target = Nothing
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
End Function


'重新查询指定窗体的数据
Public Sub RequeryForm(ByVal FormName As String, Optional ByVal Criteria As String, Optional ByVal Focus As Boolean = False)
    On Error GoTo ErrorHandler

    If Not IsFormOpen(FormName) Then Exit Sub

    Dim Target As Form
    Dim rs As DAO.Recordset

    Set Target = Forms.Item(FormName)
    If Target Is Nothing Then Exit Sub

    Target.Requery

    If Not VBA.IsMissing(Criteria) Then
        If Not StringBase.IsNullOrWhiteSpace(Criteria) Then

            Set rs = Target.RecordsetClone
            If rs Is Nothing Then Exit Sub

            If Not rs.EOF And Not rs.BOF Then
                rs.FindFirst Criteria
                If Not rs.NoMatch Then Target.Bookmark = rs.Bookmark
            End If

            rs.Close
            Set rs = Nothing
        End If

    End If

    If Focus Then Target.SetFocus
    Set Target = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Set Target = Nothing
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
End Sub

'刷新指定窗体
Public Sub RefreshForm(ByVal FormName As String, Optional ByVal Bookmark As Boolean = False, Optional ByVal Focus As Boolean = False)
    On Error GoTo ErrorHandler
    If Not IsFormOpen(FormName) Then Exit Sub

    Dim Target As Form
    Set Target = Forms.Item(FormName)
    If Target Is Nothing Then Exit Sub

    If Bookmark Then
        Dim mark As Variant
        mark = Target.Bookmark
        Target.Refresh
        Target.Bookmark = mark
    Else
        Target.Refresh
    End If

    If Focus Then Target.SetFocus
    Set Target = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Set Target = Nothing
    Exit Sub
End Sub



' --------------------------------------------------------------------------------------------------------
' -------------------------------------------- 对象检查 --------------------------------------------------
' --------------------------------------------------------------------------------------------------------

'检查某个表是否存在
Public Function TableExists(ByVal TableName As String) As Boolean
    On Error Resume Next
    If VBA.Len(TableName) > 0 Then
        Dim tdf As DAO.TableDef
        Set tdf = Application.CurrentDb.TableDefs(TableName)
        TableExists = (Not tdf Is Nothing)
        Set tdf = Nothing
    End If
    On Error GoTo 0
End Function

' 检查某个查询是否存在
Public Function QueryExists(ByVal QueryName As String) As Boolean
    On Error Resume Next
    If VBA.Len(QueryName) > 0 Then
        Dim qdf As DAO.QueryDef
        Set qdf = Application.CurrentDb.QueryDefs(QueryName)
        QueryExists = (Not qdf Is Nothing)
        Set qdf = Nothing
    End If
    On Error GoTo 0
End Function

' 检查某个窗体是否存在
Public Function FormExists(ByVal FormName As String) As Boolean
    On Error Resume Next
    If VBA.Len(FormName) > 0 Then
        Dim obj As AccessObject
        Set obj = Application.CurrentProject.AllForms(FormName)
        FormExists = (Not obj Is Nothing)
        Set obj = Nothing
    End If
    On Error GoTo 0
End Function

' 检查某个报表是否存在
Public Function ReportExists(ByVal ReportName As String) As Boolean
    On Error Resume Next
    If VBA.Len(ReportName) > 0 Then
        Dim obj As AccessObject
        Set obj = Application.CurrentProject.AllReports(ReportName)
        ReportExists = (Not obj Is Nothing)
        Set obj = Nothing
    End If
    On Error GoTo 0
End Function



' --------------------------------------------------------------------------------------------------------
' -------------------------------------------- 数据导出 --------------------------------------------------
' --------------------------------------------------------------------------------------------------------

' 导出数据表
Public Sub ExportTable(ByVal TableName As String, ByVal filename As String)
    On Error GoTo ErrorHandler
    DoCmd.Hourglass True
    Progress.Start "数据导出中......"
    If DbSql.TableCount(TableName) > 0 Then
        Progress.Run 45
        DoCmd.OutputTo acOutputTable, TableName, acFormatXLSX, Environment.DesktopPath & filename & ".xlsx", , , , acExportQualityPrint
        Progress.Run 100
        Message.Notice "已导出到系统桌面", True
    Else
        Progress.Complete
        Message.Alert "没有可导出的数据"
    End If
    DoCmd.Hourglass False
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    Call Progress.Complete
    Call Message.Error(Err)
    Exit Sub
End Sub

' 导出查询表
Public Sub ExportQuery(ByVal QueryName As String, ByVal filename As String)
    On Error GoTo ErrorHandler
    DoCmd.Hourglass True
    Progress.Start "数据导出中......"
    If DbSql.TableCount(QueryName) > 0 Then
        Progress.Run 45
        DoCmd.OutputTo acOutputQuery, QueryName, acFormatXLSX, Environment.DesktopPath & filename & ".xlsx", , , , acExportQualityPrint
        Progress.Run 100
        Message.Notice "已导出到系统桌面", True
    Else
        Progress.Complete
        Message.Alert "没有可导出的数据"
    End If
    DoCmd.Hourglass False
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    Call Progress.Complete
    Call Message.Error(Err)
    Exit Sub
End Sub

' 导出记录集
Public Function ExportRecordset(ByRef rs As DAO.Recordset, ByVal SheetName As String, Optional ByRef FieldCaptions As Collection) As Boolean
    On Error GoTo ErrorHandler
    DoCmd.Hourglass True
    Progress.Start "数据导出中......"
    If rs Is Nothing Then
        Message.Warning "无效的记录集"
        ExportRecordset = False
        Exit Function
    End If
    If rs.EOF And rs.BOF Then
        Message.Warning "记录集没有数据"
        ExportRecordset = False
        Exit Function
    End If
    Progress.Run 20
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)

    On Error Resume Next
    xlWorksheet.name = SheetName
    On Error GoTo ErrorHandler

    Progress.Run 30
    Dim i As Integer
    Dim Has As Boolean
    Dim fieldCount As Integer
    Dim captionCount As Integer

    fieldCount = rs.Fields.count
    Has = Not VBA.IsMissing(FieldCaptions)
    If Has Then captionCount = FieldCaptions.count

    Progress.Run 45
    For i = 1 To fieldCount
        If Has And i <= captionCount Then
            xlWorksheet.Cells(1, i).Value = FieldCaptions(i)
        Else
            xlWorksheet.Cells(1, i).Value = rs.Fields(i - 1).name
        End If
        xlWorksheet.Cells(1, i).Font.Bold = True
        xlWorksheet.Cells(1, i).Interior.Color = RGB(200, 200, 200)
    Next i
    xlWorksheet.Range("A2").CopyFromRecordset rs
    Progress.Run 85
    xlWorksheet.Cells.Font.size = 10
    xlWorksheet.Cells.Font.name = "宋体"
    xlWorksheet.Cells.RowHeight = 15
    If fieldCount > 0 Then xlWorksheet.Range(xlWorksheet.Cells(1, 1), xlWorksheet.Cells(1, fieldCount)).EntireColumn.AutoFit
    xlWorkbook.SaveAs Environment.DesktopPath & SheetName & ".xlsx", 51

    xlWorkbook.Close False
    xlApp.Quit
    xlApp.DisplayAlerts = True
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

    ExportRecordset = True
    DoCmd.Hourglass False
    Progress.Run 100
    Message.Notice "已导出到系统桌面", True
    Exit Function

ErrorHandler:
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlWorkbook.Close False
        xlApp.Quit
        On Error GoTo 0
    End If
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    DoCmd.Hourglass False
    Call Progress.Complete
    Message.Warning "数据导出失败"
    Call Message.Error(Err)
    ExportRecordset = False
End Function



' --------------------------------------------------------------------------------------------------------
' -------------------------------------------- 文件操作 --------------------------------------------------
' --------------------------------------------------------------------------------------------------------

' 检查某个文件是否存在
Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(FilePath) Then
        If VBA.Dir(FilePath) <> "" Then FileExists = True
    End If
    On Error GoTo 0
End Function

' 检查某个文件夹是否存在
Public Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(FolderPath) Then
        If VBA.Dir(FolderPath, VBA.vbDirectory) <> "" Then FolderExists = True
    End If
    On Error GoTo 0
End Function

' 获取文件名（含扩展名）
Public Function GetFileName(ByVal FilePath As String) As String
    GetFileName = VBA.Mid(FilePath, VBA.InStrRev(FilePath, "\") + 1)
End Function

' 获取文件名（不含扩展名）
Public Function GetFileBaseName(ByVal FilePath As String) As String
    Dim filename As String
    filename = GetFileName(FilePath)
    If VBA.InStrRev(filename, ".") > 0 Then
        GetFileBaseName = VBA.Left(filename, VBA.InStrRev(filename, ".") - 1)
    Else
        GetFileBaseName = filename
    End If
End Function

' 获取文件扩展名
Public Function GetFileExt(ByVal FilePath As String) As String
    Dim filename As String
    filename = GetFileName(FilePath)
    If VBA.InStrRev(filename, ".") > 0 Then
        GetFileExt = VBA.Mid(filename, VBA.InStrRev(filename, "."))
    Else
        GetFileExt = ""
    End If
End Function

' 获取文件父级目录
Public Function GetParentPath(ByVal FilePath As String) As String
    If VBA.Len(FilePath) > 0 And VBA.InStrRev(FilePath, "\") > 0 Then
        GetParentPath = VBA.Left(FilePath, VBA.InStrRev(FilePath, "\"))
    Else
        GetParentPath = ""
    End If
End Function

' 选择文件，返回文件路径
Public Function SelectFile(Optional ByVal DialogTitle As String, Optional ByVal Filter As String, Optional ByVal FilterTitle As String) As String
    On Error GoTo ErrorHandler

    Dim Dialog As Object
    Dim path As String

    If VBA.IsMissing(DialogTitle) Or VBA.Len(DialogTitle) = 0 Then DialogTitle = "请选择文件"
    If VBA.IsMissing(Filter) Or VBA.Len(Filter) = 0 Then Filter = "*.*"
    If VBA.IsMissing(FilterTitle) Or VBA.Len(FilterTitle) = 0 Then FilterTitle = "所有文件"

    Set Dialog = Application.FileDialog(3)
    With Dialog
        .Title = DialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add FilterTitle, Filter
        .InitialFileName = VBA.IIf(VBA.Len(Cache.SelectFolder) > 0, Cache.SelectFolder, Environment.DesktopPath)

        If .Show = -1 Then
            path = .SelectedItems(1)
        Else
            path = ""
        End If
    End With
    If VBA.Len(path) > 0 Then Cache.SelectFolder = GetParentPath(path)
    SelectFile = path

    Set Dialog = Nothing
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    SelectFile = ""
    Set Dialog = Nothing
End Function

'创建文件夹
Public Function CreateFolder(ByVal FolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    If FolderPath = "" Then Exit Function
    Dim path As String
    path = FolderPath
    If VBA.Right(path, 1) = "\" Then path = VBA.Left(path, VBA.Len(path) - 1)
    If FolderExists(path) Then Exit Function
    Dim ParentPath As String
    ParentPath = VBA.Left(path, VBA.InStrRev(path, "\") - 1)
    If ParentPath <> "" Then CreateFolder ParentPath
    CreateFolder = True
    VBA.MkDir path

    Exit Function

ErrorHandler:
    CreateFolder = False
    Call Message.Error(Err)
End Function

' 复制文件
Public Function CopyFile(ByVal SourcePath As String, ByVal DestPath As String) As Boolean
    On Error GoTo ErrorHandler
    If SourcePath = DestPath Then Exit Function
    If Not FileExists(SourcePath) Then
        Message.Warning "源文件不存在：" & SourcePath
        Exit Function
    End If

    If Not StringBase.IsNullOrEmpty(DestPath) Then
        Dim destFolder As String
        destFolder = VBA.Left(DestPath, VBA.InStrRev(DestPath, "\") - 1)
        If Not FolderExists(destFolder) Then CreateFolder destFolder
    End If

    VBA.FileCopy SourcePath, DestPath
    CopyFile = True
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
