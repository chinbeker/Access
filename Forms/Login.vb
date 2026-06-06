Option Compare Database
Option Explicit

' 主程序启动
Private Sub Form_Open(Cancel As Integer)
    App.Start
End Sub

Private Sub Form_Load()
    ' 将焦点移动到密码输入框
    If Not IsNull(Me.TextAccount.value) And Len(Me.TextAccount.value) > 0 Then
        Me.TextPassword.SetFocus
    End If
End Sub

Private Sub Form_Close()
    If Environment.Development = False And Auth.Authorized = False Then
        Application.Quit
    End If
End Sub

Private Sub BtnLogin_Click()
    On Error GoTo ErrorHandler

    If IsNull(Me!TextAccount.value) Then
        Me!TextDanger.value = "请输入账号"
    ElseIf IsNull(Me.TextPassword.value) Then
        Me!TextDanger.value = "请输入密码"
    Else
        ' 创建SQL生成器
        Dim Sql As New SqlBuilder
        Sql.Field "UserId"                                 ' 添加查询字段 UserId
        Sql.Field "UserName"                               ' 添加查询字段 UserName

        Sql.Where "Account=", DbSql.Param("Account")       ' 设置筛选条件： 字段为 Account， 筛选值：名为 Account 的SQL参数
        Sql.Where "Password=", DbSql.Param("Password")     ' 设置筛选条件： 字段为 Password，筛选值：名为 Password 的SQL参数

        Sql.Param "Account", Trim(Me.TextAccount.value)    ' 设置SQL参数 Account 的值
        Sql.Param "Password", Trim(Me.TextPassword.value)  ' 设置SQL参数 Password 的值

        Dim SqlRecordSet As DAO.Recordset
        Set SqlRecordSet = DbSql.First("Users", Sql)       ' 在 Users 表中，执行SQl

        If Not SqlRecordSet Is Nothing Then
            If SqlRecordSet.RecordCount > 0 Then

                Auth.LoginUser.UserId = SqlRecordSet.Fields.Item(0).value
                Auth.LoginUser.UserName = SqlRecordSet.Fields.Item(1).value

                SqlRecordSet.Close
                Set SqlRecordSet = Nothing

                ' 关闭SQL提示
                DoCmd.SetWarnings (False)

                ' 记住账号
                Set Sql = New SqlBuilder                                ' 创建新的SQL生成器
                Sql.Field "Account", Trim(Me.TextAccount.value)
                Sql.Where "Id=", 1
                DbSql.Update "Login", Sql                               ' 保存登录账号


                ' 更新 SQL Server 链接表密码（如果有SQL Servier链接表）
                'Dim Table As DAO.TableDef
                'Set Table = DbSql.TableDef("dbo_Settings")
                'Table.Connect = "DSN=MSSQL_ERP;Trusted_Connection=No;UID=erp;PWD=123456;APP=Microsoft Office;DATABASE=erp;Encrypt=Optional;TrustServerCertificate=Yes;"
                'Table.RefreshLink
                'Set Table = Nothing
                App.DatabaseConnected = True


                '加载设置
                Dim Setting As DAO.Recordset
                Set Setting = DbSql.First("Settings")

                App.Settings.Years = Setting.Fields("Years")
                App.Settings.Months = Setting.Fields("Months")
                App.Settings.WorkShiftId = Setting.Fields("WorkShiftId")

                ' 关闭
                Setting.Close
                Set Setting = Nothing

                Dim CurrentForm As Variant
                CurrentForm = Me.OpenArgs
                Core.CloseForm Me.name
                If Not StringBase.IsWhiteSpace(CurrentForm) Then
                    DoCmd.OpenForm CurrentForm
                Else
                    DoCmd.OpenForm "FormHome"
                End If

            Else
                Me!TextDanger.value = "账号或密码输入不正确"
            End If
        Else
            Me!TextDanger.value = "账号或密码输入不正确"
        End If
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
    If IsNull(Me.TextAccount.value) Then
        Me!TextAccount.BorderColor = 2366701
        Me!TextDanger.value = "请输入账号"
    Else
        Me!TextAccount.BorderColor = 14277081
        Me!TextDanger.value = Null
    End If
End Sub

Private Sub TextPassword_GotFocus()
    Me!TextPassword.BorderColor = 14922894
End Sub

Private Sub TextPassword_LostFocus()
    If IsNull(Me.TextPassword.value) Then
        Me!TextPassword.BorderColor = 2366701
        Me!TextDanger.value = "请输入密码"
    Else
        Me!TextPassword.BorderColor = 14277081
        Me!TextDanger.value = Null
    End If
End Sub
