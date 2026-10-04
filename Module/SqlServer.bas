Attribute VB_Name = "SqlServer"
'@Lang VBA

Option Compare Database
Option Explicit


Private DbConnection As New ADODB.Connection
Private Connected As Boolean

' ODBC
Private Const ODBC_DSN As String = "MSSQL_ERP"
Private Const ODBC_DATABASE As String = "erp"
Private Const ODBC_UID As String = "my"
Private Const ODBC_PWD As String = "123456"

' OLEDB
Private Const OLEDB_DataSource As String = "."
Private Const OLEDB_Database As String = "erp"
Private Const OLEDB_UserID As String = "my"
Private Const OLEDB_Password As String = "123456"

'其他
Private Const Timeout As Long = 3



' 连接 SQL Servier 数据库
Private Sub ConnectDatabase()
    On Error GoTo ErrorHandler
    If DbConnection Is Nothing Then Set DbConnection = New ADODB.Connection
    '如果链接关闭，则重新打开链接
    If DbConnection.State = adStateClosed Then
        With DbConnection
            .Provider = "MSOLEDBSQL.1"
            .Properties("Data Source").Value = OLEDB_DataSource
            .Properties("Initial Catalog").Value = OLEDB_Database
            .Properties("User ID").Value = OLEDB_UserID
            .Properties("Password").Value = OLEDB_Password
            .Properties("Application Name").Value = "Microsoft Access"
            '.Properties("Connection Timeout").Value = Timeout
            '.Mode = adModeReadWrite
            .ConnectionTimeout = Timeout
            .CommandTimeout = Timeout
            .Open
        End With
    End If
    Exit Sub

ErrorHandler:
    'Call Message.Warning("网络连接中断")
    Call Message.Error(Err)
    Exit Sub
End Sub

' 获取ODBC连接字符串
Private Function GetODBCConnectionString() As String
    GetODBCConnectionString = "DSN=" & ODBC_DSN & ";Trusted_Connection=No;UID=" & ODBC_UID & ";PWD=" & ODBC_PWD & ";APP=Microsoft Office;DATABASE=" & ODBC_DATABASE & ";Encrypt=Optional;TrustServerCertificate=Yes;"
End Function

' 断开连接
Public Sub Disconnect()
    On Error GoTo ErrorHandler
    DbConnection.Close
    Exit Sub
ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 更新链接表
Public Sub LinkToDatabase(ByVal TableName As String)
    Dim Table As DAO.TableDef
    If Connected = False And Core.TableExists(TableName) Then
        Set Table = DbSql.TableDef(TableName)
        Table.Connect = SqlServer.GetODBCConnectionString
        Table.RefreshLink
        Connected = True
    End If
    Set Table = Nothing
End Sub

' 创建传递查询
Public Function PassThroughQuery(Optional ByVal ReturnsRecords As Boolean = True) As DAO.QueryDef
    Dim qdf As DAO.QueryDef
    Set qdf = Application.CurrentDb.CreateQueryDef("")
    qdf.Connect = "ODBC;" & SqlServer.GetODBCConnectionString
    qdf.ReturnsRecords = ReturnsRecords
    Set PassThroughQuery = qdf
End Function

' 创建 ADODB.Command 对象 （查询命令）
Private Function CreateCommand(ByVal cmdText As String, ByVal cmdType As CommandTypeEnum) As ADODB.Command
    On Error GoTo ErrorHandler
    Call ConnectDatabase
    Set CreateCommand = New ADODB.Command
    With CreateCommand
        Set .ActiveConnection = DbConnection
        .CommandTimeout = 3
        .CommandType = cmdType
        .CommandText = cmdText
    End With
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 设置 ADODB.Parameter 参数
Public Sub SetParameter(ByRef Command As ADODB.Command, ByVal name As String, ByVal Value As Variant, ByVal DataType As DataTypeEnum, Optional ByVal size As Long, Optional ByVal Direction As ParameterDirectionEnum = adParamInput)
    On Error GoTo ErrorHandler
    If StringBase.IsNullOrEmpty(name) Then Exit Sub
    If Command Is Nothing Then Exit Sub
    Call Command.Parameters.Append(Command.CreateParameter(name, DataType, Direction, size, Value))
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 获取 ADODB.Recordset 对象（查询结果）
Private Function GetRecordSet(ByVal cmdText As String, ByVal cmdType As CommandTypeEnum) As ADODB.Recordset
    On Error GoTo ErrorHandler
    '检查查询语句（SQL字符串）是否为空，如果为空则退出
    If StringBase.IsNullOrEmpty(cmdText) Then Exit Function

    '检查链接对象是否处于打开状态，否则与数据库重现建立连接
    Call ConnectDatabase

    ' 创建 Command 对象
    Dim Command As ADODB.Command
    Set Command = CreateCommand(cmdText, cmdType)

    ' 返回 Recordset 对象（查询结果）
    Set GetRecordSet = New ADODB.Recordset
    With GetRecordSet
        Set .Source = Command
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open
    End With
    Set Command = Nothing
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 获取整个表格
Public Function Table(ByVal TableName As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsNullOrEmpty(TableName) Then Exit Function
    Set Table = GetRecordSet(TableName, adCmdTable)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 运行查询语句
Public Function query(ByVal SqlString As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsNullOrEmpty(SqlString) Then Exit Function
    Set query = GetRecordSet(SqlString, adCmdText)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 执行存储过程
Public Function StoredProc(ByVal storedProcName As String, ByRef Command As ADODB.Command) As ADODB.Recordset
    On Error GoTo ErrorHandler
    '检查存储过程名称是否为空，如果为空则退出
    If StringBase.IsNullOrEmpty(storedProcName) Then Exit Function

    '检查链接对象是否处于打开状态，否则与数据库重现建立连接
    Call ConnectDatabase

    '检查 Command 对象是否绑定已激活的 Connection 对象
    If Command.ActiveConnection Is Nothing Then Set Command.ActiveConnection = DbConnection
    With Command
        .CommandTimeout = 3
        .CommandType = adCmdStoredProc
        .CommandText = storedProcName
    End With

    ' 返回 Recordset 对象（查询结果）
    Set StoredProc = New ADODB.Recordset
    With StoredProc
        Set .Source = Command
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open
    End With

    ' 关闭 Command 对象
    Command.Cancel
    Set Command = Nothing
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
