Attribute VB_Name = "DbContext"
'@Lang VBA

Option Compare Database
Option Explicit

'检查某个表是否存在
Private Function TableExists(ByRef DB As DAO.Database, ByVal TableName As String) As Boolean
    On Error Resume Next
    If Not DB Is Nothing And VBA.Len(TableName) > 0 Then
        Dim tdf As DAO.TableDef
        Set tdf = DB.TableDefs(TableName)
        TableExists = (Not tdf Is Nothing)
        Set tdf = Nothing
    End If
    On Error GoTo 0
End Function

' 获取数据表
Public Function Table(ByRef DB As DAO.Database, ByVal TableName As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    If TableExists(DB, TableName) Then
        If ReadOnly Then
            Set Table = DB.OpenRecordset(TableName, dbOpenSnapshot)
        Else
            Set Table = DB.OpenRecordset(TableName, dbOpenDynaset, dbSeeChanges)
        End If
    End If
End Function

' 查询数据
Public Function query(ByRef DB As DAO.Database, ByVal SqlString As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    If Not DB Is Nothing Then
        If ReadOnly Then
            Set query = DB.OpenRecordset(SqlString, dbOpenSnapshot)
        Else
            Set query = DB.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
End Function

' 查询数据
Public Function Execute(ByRef DB As DAO.Database, ByVal SqlString As String) As Long
    If Not DB Is Nothing Then
        DB.Execute SqlString
        Execute = DB.RecordsAffected
    End If
End Function
