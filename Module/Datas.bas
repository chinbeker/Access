Attribute VB_Name = "Datas"
'@Lang VBA

Option Compare Database
Option Explicit


' 数据库连接池
Public Enum DatabasePoolEnum
    dpUser = 0
    dpData = 1
    dpDict = 2
    dpCache = 2
    dpLog = 2
End Enum


' 数据库配置信息
Private DB_User As DAO.Database
Private DB_Data As DAO.Database
Private DB_Dict As DAO.Database
Private DB_Cache As DAO.Database
Private DB_Log As DAO.Database

Private Const DB_USER_FILE_PATH As String = "data\user.accdb"
Private Const DB_DATA_FILE_PATH As String = "data\data.accdb"
Private Const DB_DICT_FILE_PATH As String = "data\dict.accdb"
Private Const DB_CACHE_FILE_PATH As String = "data\cache.accdb"
Private Const DB_LOG_FILE_PATH As String = "log\log.accdb"

Private Const DATA_PASSWORD As String = "********"


' 连接数据库
Private Sub ConnectDatabase()
    If DB_User Is Nothing Then Set DB_User = Application.DBEngine.OpenDatabase(Environment.CurrentPath & DB_USER_FILE_PATH, False, False, ";PWD=" & DATA_PASSWORD)
    If DB_Data Is Nothing Then Set DB_Data = Application.DBEngine.OpenDatabase(Environment.CurrentPath & DB_DATA_FILE_PATH, False, False, ";PWD=" & DATA_PASSWORD)
    If DB_Dict Is Nothing Then Set DB_Dict = Application.DBEngine.OpenDatabase(Environment.CurrentPath & DB_DICT_FILE_PATH, False, False, ";PWD=" & DATA_PASSWORD)
    If DB_Cache Is Nothing Then Set DB_Cache = Application.DBEngine.OpenDatabase(Environment.CurrentPath & DB_CACHE_FILE_PATH, False, False, ";PWD=" & DATA_PASSWORD)
    If DB_Log Is Nothing Then Set DB_Log = Application.DBEngine.OpenDatabase(Environment.CurrentPath & DB_LOG_FILE_PATH, False, False, ";PWD=" & DATA_PASSWORD)
End Sub


' 获取数据表
Public Function Table(ByVal Source As DatabasePoolEnum, ByVal TableName As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Call ConnectDatabase

    Select Case Source
        Case dpUser
            Set db = DB_User
        Case dpData
            Set db = DB_Data
        Case dpDict
            Set db = DB_Dict
        Case dpCache
            Set db = DB_Cache
        Case dpLog
            Set db = DB_Log
        Case Else
            Set db = DB_Data
    End Select

    On Error Resume Next
    Set tdf = db.TableDefs(TableName)
    On Error GoTo 0
    If tdf Is Nothing Then
        Err.Raise vbObjectError + 1000, "Data.Table", "Table not found: " & TableName
    End If

    If ReadOnly Then
        Set Table = db.OpenRecordset(TableName, dbOpenSnapshot)
    Else
        Set Table = db.OpenRecordset(TableName, dbOpenDynaset, dbSeeChanges)
    End If

    Set db = Nothing
    Set tdf = Nothing
End Function
