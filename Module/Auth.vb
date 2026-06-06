Option Compare Database
Option Explicit

Public LoginUser As New User

' 判断是否是登录授权
Public Function Authorized() As Boolean

    '检查全局变量是否失效
    If App.AppStarted And App.DatabaseConnected = False Then
        Dim AppConfig As DAO.Recordset
        Set AppConfig = DbSql.First("Application")
        If Not AppConfig Is Nothing Then
            App.AppName = AppConfig.Fields("AppName")
            App.AppTitle = AppConfig.Fields("AppTitle")
            App.Installed = AppConfig.Fields("Installed")
            Environment.Development = AppConfig.Fields("Development")
            App.DatabaseConnected = True
        End If
        AppConfig.Close
        Set AppConfig = Nothing
    End If

    If App.AppStarted = False Or Auth.LoginUser.UserId = 0 Then
        Authorized = False
    Else
        Authorized = True
    End If

End Function

'无授权，进入登录页面
Public Sub Login(Optional ByVal FormName As String)
    If Not IsMissing(FormName) And Not StringBase.IsWhiteSpace(FormName) Then
        DoCmd.OpenForm "FormLogin", , , , , , FormName
    Else
        DoCmd.OpenForm "FormLogin"
    End If
End Sub
