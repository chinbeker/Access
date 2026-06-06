Option Compare Database
Option Explicit

'窗体打开
Private Sub Form_Open(Cancel As Integer)
    If Auth.Authorized = False Then
        Cancel = True
        Auth.Login Me.name
    Else
        Me.RecordSource = "Settings"
        Me.Filter = ""
        Me.FilterOn = False
        Me.OrderByOn = False
        Me.OrderBy = ""
    End If
End Sub

'窗体加载
Private Sub Form_Load()
    Me.UserName.value = Auth.LoginUser.UserName
End Sub
