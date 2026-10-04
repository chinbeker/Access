Attribute VB_Name = "DbPool"
'@Lang VBA

Option Compare Database
Option Explicit

' user
Private DB_User As DAO.Database


Private Function GetDatabasePath(ByVal path As String) As String
    path = VBA.Trim(path)

    If StringBase.IsNullOrEmpty(path) Then
        Err.Raise vbObjectError + 1001, "DbPool.GetDatabasePath", "数据库名称不能为空"
        Exit Function
    Else
        path = Environment.CurrentPath & path
        path = VBA.Replace(path, "\\", "\")
        If Core.FileExists(path) Then
            GetDatabasePath = path
        Else
            Err.Raise 53, "DbPool.GetDatabasePath", "数据库文件 " & path & " 不存在"
            Exit Function
        End If
    End If
End Function


' 数据库 (User)
Public Function DbUser() As DAO.Database
    On Error GoTo ErrorHandler
    If DB_User Is Nothing Then Set DB_User = Application.DBEngine.OpenDatabase(GetDatabasePath("data\user.accdb"), False, False, ";PWD=123456")
    Set DbUser = DB_User
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
