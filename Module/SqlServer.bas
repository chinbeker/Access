Attribute VB_Name = "SqlServer"
'@Lang VBA

Option Compare Database
Option Explicit

'Database Connect
Private DbConnection As New ADODB.Connection

'Sql Server
Private Sub CreateConnection()
    '如果链接关闭，则重新打开链接
    On Error GoTo ErrorHandler
    If DbConnection.State = 0 Then
        With DbConnection
            .Provider = "MSOLEDBSQL.1"

            ' 方案一、使用系统DSN数据源
            '.Properties("Data Source").Value = "ERP"

            ' 方案二、使用数据库实例名称
            '.Properties("Data Source").Value = "COMPUTERNAME\EXPRESS"

            ' 方案三、使用IP地址
            '.Properties("Network Address").Value = "192.168.101.235"

            ' 方案四、使用本机
            .Properties("Data Source").value = "."

            .Properties("Initial Catalog").value = "erp"
            .Properties("User ID").value = "users"
            .Properties("Password").value = "123456"
            .Properties("Application Name").value = "Microsoft Access"
            '.Properties("Connection Timeout").Value = 30
            .ConnectionTimeout = 3
            .Open
        End With
    End If
    Exit Sub

ErrorHandler:
    Call Message.Warning("网络连接中断")
End Sub

' 关闭 DbConnection 连接
Public Sub CloseConnection()
    On Error GoTo ErrorHandler
    DbConnection.Close
    Exit Sub
ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub


' 创建 ADODB.Command 对象 （查询命令）
Private Function CreateCommand(ByVal cmdText As String, ByVal cmdType As CommandTypeEnum) As ADODB.Command
    On Error GoTo ErrorHandler
    Call CreateConnection
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
Public Sub SetParameter(ByRef Command As ADODB.Command, ByVal name As String, ByVal value As Variant, ByVal DataType As DataTypeEnum, Optional ByVal Size As Long, Optional ByVal Direction As ParameterDirectionEnum = adParamInput)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(name) Then Exit Sub
    If Command Is Nothing Then Exit Sub
    Call Command.Parameters.Append(Command.CreateParameter(name, DataType, Direction, Size, value))
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

' 获取 ADODB.Recordset 对象（查询结果）
Private Function GetRecordSet(ByVal cmdText As String, ByVal cmdType As CommandTypeEnum) As ADODB.Recordset
    On Error GoTo ErrorHandler
    '检查查询语句（SQL字符串）是否为空，如果为空则退出
    If StringBase.IsWhiteSpace(cmdText) Then Exit Function

    '检查链接对象是否处于打开状态，否则与数据库重现建立连接
    Call CreateConnection

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
    If StringBase.IsWhiteSpace(TableName) Then Exit Function
    Set Table = GetRecordSet(TableName, adCmdTable)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 运行查询语句
Public Function Query(ByVal SqlString As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(SqlString) Then Exit Function
    Set Query = GetRecordSet(SqlString, adCmdText)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 执行存储过程
Public Function StoredProc(ByVal storedProcName As String, ByRef Command As ADODB.Command) As ADODB.Recordset
    On Error GoTo ErrorHandler
    '检查存储过程名称是否为空，如果为空则退出
    If StringBase.IsWhiteSpace(storedProcName) Then Exit Function

    '检查链接对象是否处于打开状态，否则与数据库重现建立连接
    Call CreateConnection

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
