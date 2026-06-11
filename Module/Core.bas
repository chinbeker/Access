Attribute VB_Name = "Core"
'@Lang VBA

Option Compare Database
Option Explicit


'打开指定窗体（统一验证授权）
Public Sub OpenForm(ByVal FormName As String, Optional ByVal Condition As String, Optional ByVal OpenArgs As String)
    If StringBase.IsWhiteSpace(FormName) Then Exit Sub
    If Auth.Authorized = False Then
        Call Auth.Login(FormName)
    Else
        DoCmd.OpenForm FormName, , , Condition, , , OpenArgs
    End If
End Sub

' 关闭指定窗体
Public Sub CloseForm(ByVal FormName As String)
    If StringBase.IsWhiteSpace(FormName) Then Exit Sub
    DoCmd.Close acForm, FormName, acSaveNo
End Sub

'刷新指定窗体的数据
Public Sub RequeryForm(ByVal FormName As String)
    If StringBase.IsWhiteSpace(FormName) Then Exit Sub
    Forms.Item(FormName).Requery
End Sub

' 导出数据表
Public Sub ExportTable(ByVal TableName As String, ByVal FileName As String)
    If DbSql.TableCount(TableName) > 0 Then
        DoCmd.OutputTo acOutputTable, TableName, acFormatXLSX, Environment.DesktopPath & FileName & ".xlsx", , , , acExportQualityPrint
        Message.Alert "已导出到系统桌面"
    Else
        Message.Alert "没有可导出的数据"
    End If
End Sub

' 导出查询表
Public Sub ExportQuery(ByVal QueryName As String, ByVal FileName As String)
    If DbSql.TableCount(QueryName) > 0 Then
        DoCmd.OutputTo acOutputQuery, QueryName, acFormatXLSX, Environment.DesktopPath & FileName & ".xlsx", , , , acExportQualityPrint
        Message.Alert "已导出到系统桌面"
    Else
        Message.Alert "没有可导出的数据"
    End If
End Sub
