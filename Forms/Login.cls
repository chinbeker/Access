VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Lang VBA

Option Compare Database
Option Explicit

' 主程序启动
Private Sub Form_Open(Cancel As Integer)
    App.Start
End Sub

' 将焦点移动到密码输入框
Private Sub Form_Load()
    If Not VBA.IsNull(Me.TextAccount.Value) And VBA.Len(Me.TextAccount.Value) > 0 Then
        Me.TextPassword.SetFocus
    End If
End Sub

' 退出登录
Private Sub Form_Close()
    If App.Started And Auth.Authenticated = False And Environment.Development = False Then
        Application.Quit
    End If
End Sub

' 登录按钮
Private Sub BtnLogin_Click()
    On Error GoTo ErrorHandler

    If IsNull(Me!TextAccount.Value) Then
        Me!TextDanger.Value = "请输入账号"
    ElseIf IsNull(Me.TextPassword.Value) Then
        Me!TextDanger.Value = "请输入密码"
    Else
        '-------------------------------------------------------------------------------------------------------------
        '-------------------------------------------- 使用 Access 数据库 ---------------------------------------------
        '-------------------------------------------------------------------------------------------------------------
        Dim User As DAO.Recordset
        Dim Sql As New SqlBuilder                                  ' 创建SQL生成器

        Sql.Field "UserId"                                         ' 添加查询字段 UserId
        Sql.Field "UserName"                                       ' 添加查询字段 UserName

        Sql.From "Users"                                           ' 设置数据源(表)

        Sql.Where "Account=", DbSql.Parameter("Account")           ' 设置筛选条件： 字段 Account= ， 筛选值：名为 Account 的SQL参数
        Sql.Where "Password=", DbSql.Parameter("Password")         ' 设置筛选条件： 字段 Password= ，筛选值：名为 Password 的SQL参数

        Sql.Param "Account", VBA.Trim(Me.TextAccount.Value)                         ' 设置SQL参数 Account 的值
        Sql.Param "Password", Security.SHA256(VBA.Trim(Me.TextPassword.Value))      ' 设置SQL参数 Password 的值 (SHA256加密)

        DbSql.UseCustomDatabase DbPool.DbUser                      ' 使用指定数据库（user数据库保存着用户账户密码，不使用链接表）
        Set User = DbSql.First(Sql)                                ' 执行SQl，返回第1条记录

        If Auth.Login(User) Then
            ' 记住账号
            Set Sql = New SqlBuilder
            Sql.Field "Account", VBA.Trim(Me.TextAccount.Value)
            Sql.Where "Id=", 1
            DbSql.Update "Login", Sql
            Set Sql = Nothing

            ' 进入首页
            Dim SourceForm As Variant
            SourceForm = Me.OpenArgs
            Core.CloseForm Me.name
            If Not StringBase.IsNullOrEmpty(SourceForm) Then
                DoCmd.OpenForm SourceForm
            Else
                DoCmd.OpenForm "FormHome"
            End If

        '登录失败
        Else
            Me!TextDanger.Value = "账号或密码输入不正确"
        End If
        '-------------------------------------------------------------------------------------------------------------


        '-------------------------------------------------------------------------------------------------------------
        '-------------------------------------------- 使用SQL Server 数据库 -------------------------------------------
        '-------------------------------------------------------------------------------------------------------------
        'Dim SqlCommand As New ADODB.Command
        'Dim User As ADODB.Recordset
        ''设置参数
        'SqlServer.SetParameter SqlCommand, "Account", VBA.Trim(Me.TextAccount.value), adVarChar, 30
        'SqlServer.SetParameter SqlCommand, "Password", Security.SHA256(VBA.Trim(Me.TextPassword.value)), adVarChar, 64
        '执行存储过程
        'Set User = SqlServer.StoredProc("dbu.UserLogin", SqlCommand)

        'If Auth.LoginFromSqlServer(User) Then
        '    ' 记住账号
        '    Dim Sql As New SqlBuilder
        '    Sql.Field "Account", VBA.Trim(Me.TextAccount.value)
        '    Sql.Where "Id=", 1
        '    DbSql.Update "Login", Sql
        '    Set Sql = Nothing

        '    ' 进入首页
        '    Dim SourceForm As Variant
        '    SourceForm = Me.OpenArgs
        '    Core.CloseForm Me.name
        '    If Not StringBase.IsNullOrEmpty(SourceForm) Then
        '        DoCmd.OpenForm SourceForm
        '    Else
        '        DoCmd.OpenForm "FormHome"
        '    End If
        ''登录失败
        'Else
        '    Me!TextDanger.value = "账号或密码输入不正确"
        'End If
        '-------------------------------------------------------------------------------------------------------------

        Set User = Nothing
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Sub TextAccount_GotFocus()
    Me!TextAccount.BorderColor = 14922894
End Sub

Private Sub TextAccount_LostFocus()
    If IsNull(Me.TextAccount.Value) Then
        Me!TextAccount.BorderColor = 2366701
        Me!TextDanger.Value = "请输入账号"
    Else
        Me!TextAccount.BorderColor = 14277081
        Me!TextDanger.Value = Null
    End If
End Sub

Private Sub TextPassword_GotFocus()
    Me!TextPassword.BorderColor = 14922894
End Sub
Private Sub TextPassword_LostFocus()
    If IsNull(Me.TextPassword.Value) Then
        Me!TextPassword.BorderColor = 2366701
        Me!TextDanger.Value = "请输入密码"
    Else
        Me!TextPassword.BorderColor = 14277081
        Me!TextDanger.Value = Null
    End If
End Sub
