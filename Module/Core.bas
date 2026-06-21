Attribute VB_Name = "Core"
'@Lang VBA

Option Compare Database
Option Explicit


' 窗体是否已打开
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    If Not StringBase.IsNullOrEmpty(FormName) Then
        IsFormOpen = (Application.SysCmd(acSysCmdGetObjectState, acForm, FormName) <> 0)
    Else
        IsFormOpen = False
    End If
End Function

'打开指定窗体（统一验证授权）
Public Sub OpenForm(ByVal FormName As String, Optional ByVal Condition As String, Optional ByVal OpenArgs As String)
    On Error GoTo ErrorHandler
    If StringBase.IsNullOrEmpty(FormName) Then Exit Sub
    If Not Auth.Authorized Then
        Call Auth.RedirectToLogin(FormName)
    Else
        DoCmd.OpenForm FormName, , , Condition, , , OpenArgs
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
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

'刷新指定窗体的数据
Public Sub RequeryForm(ByVal FormName As String)
    On Error GoTo ErrorHandler
    If IsFormOpen(FormName) Then Forms.Item(FormName).Requery
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 导出数据表
Public Sub ExportTable(ByVal TableName As String, ByVal FileName As String)
    On Error GoTo ErrorHandler
    If DbSql.TableCount(TableName) > 0 Then
        DoCmd.OutputTo acOutputTable, TableName, acFormatXLSX, Environment.DesktopPath & FileName & ".xlsx", , , , acExportQualityPrint
        Message.Alert "已导出到系统桌面"
    Else
        Message.Alert "没有可导出的数据"
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 导出查询表
Public Sub ExportQuery(ByVal QueryName As String, ByVal FileName As String)
    On Error GoTo ErrorHandler
    If DbSql.TableCount(QueryName) > 0 Then
        DoCmd.OutputTo acOutputQuery, QueryName, acFormatXLSX, Environment.DesktopPath & FileName & ".xlsx", , , , acExportQualityPrint
        Message.Alert "已导出到系统桌面"
    Else
        Message.Alert "没有可导出的数据"
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub


'检查某个表是否存在
Public Function TableExists(ByVal TableName As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(TableName) Then
        Dim tdf As DAO.TableDef
        Set tdf = Application.CurrentDb.TableDefs(TableName)
        If Not tdf Is Nothing Then
            TableExists = (Left(tdf.Name, 4) <> "MSys")
            Set tdf = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 检查某个查询是否存在
Public Function QueryExists(ByVal QueryName As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(QueryName) Then
        Dim qdf As DAO.QueryDef
        Set qdf = Application.CurrentDb.QueryDefs(QueryName)
        QueryExists = Not qdf Is Nothing
        Set qdf = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 检查某个窗体是否存在
Public Function FormExists(ByVal FormName As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(FormName) Then
        Dim obj As AccessObject
        Set obj = Application.CurrentProject.AllForms(FormName)
        FormExists = Not obj Is Nothing
        Set obj = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 检查某个报表是否存在
Public Function ReportExists(ByVal ReportName As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(ReportName) Then
        Dim obj As AccessObject
        Set obj = Application.CurrentProject.AllReports(ReportName)
        ReportExists = Not obj Is Nothing
        Set obj = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Function
