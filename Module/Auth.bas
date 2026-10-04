Attribute VB_Name = "Auth"
'@Lang VBA

Option Compare Database
Option Explicit

Public Connected As Boolean                '连接状态


' 已登录
Public Function Authenticated() As Boolean
    If State.GetValue("Authenticated") = True Then Authenticated = True
End Function

' 授权
Public Function Authorized() As Boolean
    '检查用户认证
    If App.Started = True And Auth.Authenticated = True Then
        '检查全局变量是否失效
        If Not Auth.Connected Then Auth.Connected = True
        If Token.UserId <> 0 Then Authorized = True
    End If
End Function

'无授权，进入登录页面
Public Sub RedirectToLogin(Optional ByVal FormName As String)
    If VBA.Len(FormName) > 0 Then
        DoCmd.OpenForm "USysFormLogin", , , , , , FormName
    Else
        DoCmd.OpenForm "USysFormLogin"
    End If
End Sub


' 保存登录信息
Private Function SetUserToken(ByRef User As DAO.Recordset) As Boolean
    If User.Fields.count > 0 Then

        ' 要保存的用户信息
        Token.SetValue "UserId", User.Fields.Item(0).Value
        Token.SetValue "UserName", User.Fields.Item(1).Value
        '......

        SetUserToken = True
        User.Close
        Set User = Nothing
    End If
End Function
Private Function SetUserTokenFromSqlServer(ByRef User As ADODB.Recordset) As Boolean
    If User.Fields.count > 0 Then

        ' 要保存的用户信息
        Token.SetValue "UserId", User.Fields.Item(0).Value
        Token.SetValue "UserName", User.Fields.Item(1).Value
        '......

        SetUserTokenFromSqlServer = True
        User.Close
        Set User = Nothing
    End If
End Function



' 加载用户设置
Private Function LoadUserSetting() As Boolean
    Dim tSetting As DAO.Recordset
    If Bootstrap.UseSqlServer = True Then Call SqlServer.LinkToDatabase("dbo_Settings")
    DbSql.UseDefaultDatabase

    Set tSetting = DbSql.TableFirst("dbo_Settings")
    If Not tSetting Is Nothing Then
        If tSetting.RecordCount > 0 Then

            ' 要读取的用户设置信息
            Setting.SetValue "Years", tSetting.Fields("Years").Value
            Setting.SetValue "LunarYear", tSetting.Fields("LunarYear").Value
            Setting.SetValue "Months", tSetting.Fields("Months").GetValu
            ' ......

            LoadUserSetting = True
        End If
        tSetting.Close
    End If
    Set tSetting = Nothing
End Function

' 登录
Public Function Login(ByRef User As DAO.Recordset) As Boolean
    If Not User Is Nothing Then
        If User.RecordCount > 0 Then
            If SetUserToken(User) And LoadUserSetting() Then
                State.SetValue "Authenticated", True
                Auth.Connected = True
                Login = True
                Set User = Nothing
            End If
        End If
    Else
        Err.Raise vbObjectError + 1001, "Auth.Login", "系统错误"
    End If
End Function



' 登录
Public Function LoginFromSqlServer(ByRef User As ADODB.Recordset) As Boolean
    If Not User Is Nothing Then
        If User.RecordCount > 0 Then
            If SetUserTokenFromSqlServer(User) And LoadUserSetting() Then
                State.SetValue "Authenticated", True
                Auth.Connected = True
                Call SqlServer.Disconnect
                LoginFromSqlServer = True
                Set User = Nothing
            End If
        End If
    Else
        Err.Raise vbObjectError + 1001, "Auth.Login", "系统错误"
    End If
End Function

' 退出登录
Public Sub Logout()
    State.SetValue "Authenticated", False
End Sub
