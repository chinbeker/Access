Option Compare Database
Option Explicit

'数据库对象
Private Database As DAO.Database

' 建立数据库连接
Private Sub CreateConnection()
    On Error GoTo ErrorHandler
    If Database Is Nothing Then Set Database = CurrentDb()
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 引用字段（拼接SQL字符串时不会加引号）
Public Function Field(ByVal Name As String) As String
    If Not StringBase.IsWhiteSpace(Name) Then Field = "[$$]" & Name
End Function

' 引用表达式（拼接SQL字符串时不会加引号）
Public Function Expression(ByVal expr As String) As String
    Expression = DbSql.Field(expr)
End Function
' 引用参数（拼接SQL字符串时不会加引号）
Public Function Param(ByVal Name As String) As String
    If Not StringBase.IsWhiteSpace(Name) Then Param = "[$$][Param_" & Name & "]"
End Function


' 运行任意查询类SQL语句（动态集）
Public Function SelectDynaset(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(SqlString) Then Exit Function
    Call CreateConnection
    Set SelectDynaset = DbSql.Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 运行任意查询类SQL语句（快照）
Public Function SelectSnapshot(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(SqlString) Then Exit Function
    Call CreateConnection
    Set SelectSnapshot = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 运行任意非查询SQL语句
Public Function Execute(ByVal SqlString As String) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(SqlString) Then Exit Function
    Call CreateConnection
    DbSql.Database.Execute SqlString
    Execute = DbSql.Database.RecordsAffected
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
' 获取表定义
Public Function TableDef(ByVal Name As String) As DAO.TableDef
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Name) Then Exit Function
    Call CreateConnection
    Set TableDef = DbSql.Database.TableDefs(Name)
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 获取整张表（本地表）
Public Function OpenTable(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Name) Then Exit Function
    Call CreateConnection
    Set OpenTable = DbSql.Database.OpenRecordset(Name, dbOpenTable)
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 获取整张表（动态集）
Public Function TableDynaset(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Name) Then Exit Function
    Call CreateConnection
    Set TableDynaset = DbSql.Database.OpenRecordset(Name, dbOpenDynaset, dbSeeChanges)
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 获取整张表（快照）
Public Function TableSnapshot(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Name) Then Exit Function
    Call CreateConnection
    Set TableSnapshot = DbSql.Database.OpenRecordset(Name, dbOpenSnapshot)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 插入数据
Public Function Insert(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function

    Sql.Into TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(4)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Def.Execute dbFailOnError
            Insert = Def.RecordsAffected
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            DbSql.Database.Execute SqlString
            Insert = DbSql.Database.RecordsAffected
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 更新数据
Public Function Update(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(3)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Def.Execute dbFailOnError
            Update = Def.RecordsAffected
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            DbSql.Database.Execute SqlString
            Update = DbSql.Database.RecordsAffected
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 删除数据
Public Function Delete(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler

    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(2)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Def.Execute dbFailOnError
            Delete = Def.RecordsAffected
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            DbSql.Database.Execute SqlString
            Delete = DbSql.Database.RecordsAffected
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 清空数据
Public Function Clear(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Call CreateConnection
    DbSql.Database.Execute "DELETE FROM " & TableName
    Clear = DbSql.Database.RecordsAffected
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 统计数量
Public Function Count(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(1)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        Dim rs As DAO.Recordset
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set rs = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set rs = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If

        Set Sql = Nothing
        Count = rs(0)
        rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 统计整张表格数量
Public Function TableCount(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.From TableName
    TableCount = DbSql.Count(TableName, Sql)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 返回记录（快照）
Public Function Find(ByVal TableName As String, ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set Find = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set Find = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 返回记录（动态集）
Public Function Record(ByVal TableName As String, ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set Record = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set Record = DbSql.Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 第一条记录（快照）
Public Function First(ByVal TableName As String, Optional ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If IsMissing(Sql) Or Sql Is Nothing Then
        Set Sql = New SqlBuilder
        Sql.SelectAll
    End If
    Sql.Top 1
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set First = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set First = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 第一条记录（动态集）
Public Function FirstRecord(ByVal TableName As String, Optional ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If IsMissing(Sql) Or Sql Is Nothing Then
        Set Sql = New SqlBuilder
        Sql.SelectAll
    End If
    Sql.Top 1
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set FirstRecord = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set FirstRecord = DbSql.Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 表格第一条记录（快照）
Public Function TableFirst(ByVal TableName As String, ByVal OrderField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(OrderField) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.Top 1
    Sql.SelectAll
    Sql.From TableName
    Sql.Order OrderField
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set TableFirst = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set TableFirst = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 表格第一条记录（动态集）
Public Function TableFirstRecord(ByVal TableName As String, ByVal OrderField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(OrderField) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.Top 1
    Sql.From TableName
    Sql.Order OrderField
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set TableFirstRecord = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set TableFirstRecord = DbSql.Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 表格最后一条记录（快照）
Public Function TableLast(ByVal TableName As String, ByVal OrderField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(OrderField) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.Top 1
    Sql.From TableName
    Sql.Order OrderField, True
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set TableLast = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set TableLast = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 表格最后一条记录（动态集）
Public Function TableLastRecord(ByVal TableName As String, ByVal OrderField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(OrderField) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.Top 1
    Sql.From TableName
    Sql.Order OrderField, True
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set TableLastRecord = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set TableLastRecord = DbSql.Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 获取第一条记录的指定字段值
Public Function GetValue(ByVal TableName As String, ByRef Sql As SqlBuilder) As Variant
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Sql.Top 1
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    If Len(SqlString) > 0 Then
        Call CreateConnection
        Dim rs As DAO.Recordset

        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Set rs = Def.OpenRecordset(dbOpenSnapshot)
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            Set rs = Database.OpenRecordset(SqlString, dbOpenSnapshot)
        End If

        If Not rs.EOF Then
            GetValue = rs(0)
        Else
            GetValue = Null
        End If

        rs.Close
        Set rs = Nothing
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 获取第一条记录的指定字段值
Public Function GetValueFromSql(ByVal SqlString As String) As Variant
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(SqlString) Then Exit Function

    Call CreateConnection
    Dim rs As DAO.Recordset
    Set rs = DbSql.Database.OpenRecordset(SqlString, dbOpenSnapshot)

    If Not rs.EOF Then
        GetValueFromSql = rs(0)
    Else
        GetValueFromSql = Null
    End If

    rs.Close
    Set rs = Nothing

    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 快速查询记录指定字段值
Public Function Lookup(ByVal TableName As String, ByVal Field As String, ByVal Condition As String, Optional ByVal OrderField As String) As Variant
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(Field) Then Exit Function
    If StringBase.IsWhiteSpace(Condition) Then Exit Function
    If IsMissing(OrderField) Or StringBase.IsWhiteSpace(OrderField) Then
        Lookup = Application.DLookup(Field, TableName, Condition)
    Else
        Dim Sql As New SqlBuilder
        Sql.Top 1
        Sql.Field Field
        Sql.From TableName
        Sql.Where Condition
        Sql.Order OrderField
        Dim SqlString As String
        SqlString = Sql.ToSqlString(0)
        Set Sql = Nothing
        If Len(SqlString) > 0 Then
            Lookup = Application.DLookup(Field, SqlString)
        Else
            Lookup = Null
        End If
    End If
    Exit Function
ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 快速查询第一条记录指定字段值
Public Function FirstValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderField As String) As Variant
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(Field) Then Exit Function
    If IsMissing(OrderField) Or StringBase.IsWhiteSpace(OrderField) Then
        If IsMissing(Condition) Or StringBase.IsWhiteSpace(Condition) Then
            FirstValue = Application.DFirst(Field, TableName)
        Else
            FirstValue = Application.DFirst(Field, TableName, Condition)
        End If
    Else
        Dim Sql As New SqlBuilder
        Sql.Top 1
        Sql.Field Field
        Sql.From TableName
        Sql.Where Condition
        Sql.Order OrderField
        Dim SqlString As String
        SqlString = Sql.ToSqlString(0)
        Set Sql = Nothing
        If Len(SqlString) > 0 Then
            FirstValue = Application.DFirst(Field, SqlString)
        Else
            FirstValue = Null
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 快速查询最后一条记录指定字段值
Public Function LastValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderField As String) As Variant
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(Field) Then Exit Function
    If IsMissing(OrderField) Or StringBase.IsWhiteSpace(OrderField) Then
        If IsMissing(Condition) Or StringBase.IsWhiteSpace(Condition) Then
            LastValue = Application.DLast(Field, TableName)
        Else
            LastValue = Application.DLast(Field, TableName, Condition)
        End If
    Else
        Dim Sql As New SqlBuilder
        Sql.Top 1
        Sql.Field Field
        Sql.From TableName
        Sql.Where Condition
        Sql.Order OrderField, True
        Dim SqlString As String
        SqlString = Sql.ToSqlString(0)
        Set Sql = Nothing
        If Len(SqlString) > 0 Then
            LastValue = Application.DFirst(Field, SqlString)
        Else
            LastValue = Null
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 设置第一条记录的指定字段值
Public Function SetValue(ByVal TableName As String, ByVal Field As String, ByVal value As Variant, ByVal Condition As String) As Boolean
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    If StringBase.IsWhiteSpace(Field) Then Exit Function
    If StringBase.IsWhiteSpace(Condition) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.Top 1
    Sql.From TableName
    Sql.Field Field, value
    Sql.Where Condition
    Dim SqlString As String
    SqlString = Sql.ToSqlString(3)

    If Len(SqlString) > 0 Then
        Call CreateConnection
        Dim Affected As Long

        If Sql.HasParam Then
            Dim Def As DAO.QueryDef
            Set Def = DbSql.Database.CreateQueryDef("", SqlString)
            Sql.SetParam Def
            Set Sql = Nothing
            Def.Execute dbFailOnError
            Affected = Def.RecordsAffected
            Def.Close
            Set Def = Nothing
        Else
            Set Sql = Nothing
            DbSql.Database.Execute SqlString
            Affected = DbSql.Database.RecordsAffected
        End If

        If Affected > 0 Then
            SetValue = True
        Else
            SetValue = False
        End If

    Else
        SetValue = False
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
