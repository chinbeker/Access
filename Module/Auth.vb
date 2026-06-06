Option Compare Database
Option Explicit

Public LoginUser As New User

' 判断是否是登录授权
Public Function Authorized() As Boolean
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
