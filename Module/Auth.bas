Attribute VB_Name = "Auth"
'@Lang VBA

Option Compare Database
Option Explicit

Public Token As New UserToken
Public UserId As Long

' 已登录
Public Function Authenticated() As Boolean
    If State.GetValue("Authenticated") = True Then Authenticated = True
End Function

' 授权
Public Function Authorized() As Boolean
    '检查用户认证
    If App.Started And Auth.Authenticated Then
        '检查全局变量是否失效
        If Not App.Connected Then
            App.Initialize
            App.Connected = True
            If Auth.Token Is Nothing Then Set Auth.Token = New UserToken
            Auth.UserId = Auth.Token.UserId
        End If
        If Auth.UserId <> 0 Then Authorized = True
    End If
End Function

'无授权，进入登录页面
Public Sub RedirectToLogin(Optional ByVal FormName As String)
    If Not StringBase.IsNullOrEmpty(FormName) Then
        DoCmd.OpenForm "FormLogin", , , , , , FormName
    Else
        DoCmd.OpenForm "FormLogin"
    End If
End Sub

' 登录认证
Public Sub LoginSuccessful()
    Auth.UserId = Auth.Token.UserId
    State.SetValue "Authenticated", True
    App.Connected = True
End Sub
