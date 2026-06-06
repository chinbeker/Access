Option Compare Database
Option Explicit


Public AppName As String             '应用名称
Public AppTitle As String            '标题名称
Public Installed As Boolean          '安装状态
Public AppStarted As Boolean         '启动状态
Public DatabaseConnected As Boolean  '数据库连接状态

Public Settings As New Setting       '设置

' 应用启动
Public Sub Start()
    If App.AppStarted Then Exit Sub

    '禁用执行 Sql 时的弹窗提示
    DoCmd.SetWarnings (False)

    '读取配置
    Dim AppConfig As DAO.Recordset
    Set AppConfig = DbSql.FirstRecord("Application")

    If Not AppConfig Is Nothing Then
        App.AppName = AppConfig.Fields("AppName")
        App.AppTitle = AppConfig.Fields("AppTitle")
        App.Installed = AppConfig.Fields("Installed")
        Environment.Development = AppConfig.Fields("Development")
    Else
        Call Message.Error(Err)
        Exit Sub
    End If


    '生产环境执行代码

    '首次安装
    If Environment.Development = False And App.Installed = False Then
        '信任安全位置
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")

        Dim path As String
        path = "HKCU\SOFTWARE\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location8\"

        Dim value As String
        Dim IsDriverInstalled As Boolean

        On Error Resume Next
        value = wsh.RegRead(path)
        IsDriverInstalled = (Err.Number = 0)

        On Error GoTo 0
        If Not IsDriverInstalled Then
            wsh.RegWrite path & "Path", Environment.CurrentPath, "REG_SZ"
            wsh.RegWrite path & "AllowSubfolders", 1, "REG_DWORD"
            wsh.RegWrite path & "Description", "", "REG_SZ"
            wsh.RegWrite path & "Date", "02/01/2026 12:00", "REG_SZ"

            '安装 ODBC_Driver_18_for_SQL_Server 驱动
            value = ""
            IsDriverInstalled = False

            On Error Resume Next
            value = wsh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\ODBC Driver 18 for SQL Server")
            IsDriverInstalled = (Err.Number = 0)

            On Error GoTo 0
            If Not IsDriverInstalled Then
                path = Environment.CurrentPath & "resource\driver\Microsoft_ODBC_Driver_18_for_SQL_Server_18.6.1.1_x64.msi"
                If VBA.Dir(path) = "" Then
                    Message.Warning "未找到 Microsoft_ODBC_Driver_18_for_SQL_Server 驱动文件"
                    Exit Sub
                End If
                wsh.Run "msiexec.exe /i """ & path & """ /passive /norestart", 0, True
            End If

            '安装 OLE_DB_Driver_18_for_SQL_Server 驱动
            value = ""
            IsDriverInstalled = False

            On Error Resume Next
            value = wsh.RegRead("HKLM\SOFTWARE\Microsoft\MSOLEDBSQL\InstalledVersion")
            IsDriverInstalled = (Err.Number = 0)

            On Error GoTo 0
            If Not IsDriverInstalled Or value <> "18.7.5.0" Then
                path = Environment.CurrentPath & "resource\driver\Microsoft_OLE_DB_Driver_18_for_SQL_Server_18.7.5_x64.msi"
                If VBA.Dir(path) = "" Then
                    Message.Warning "未找到 Microsoft_OLE_DB_Driver_18_for_SQL_Server 驱动文件"
                    Exit Sub
                End If
                wsh.Run "msiexec.exe /i """ & path & """ /passive /norestart", 0, True
            End If

            ' 导入 ODBC 数据源配置
            path = Environment.CurrentPath & "config\ODBC_Data_Sources.reg"
            If VBA.Dir(path) = "" Then
                Message.Warning "未找到 ODBC 数据源配置文件"
                Exit Sub
            End If

            On Error Resume Next
            wsh.Run "regedit.exe /s """ & path & """", 0, True

            If Err.Number <> 0 Then
                Message.Warning "配置文件导入失败，请联系管理员"
                Exit Sub
            End If

            '更新链接表目录
            Dim tdf As DAO.TableDef
            For Each tdf In CurrentDb.TableDefs
                If Len(tdf.Connect) > 0 Then
                    If (tdf.Attributes And dbAttachedTable) <> 0 And (tdf.Attributes And dbAttachedODBC) = 0 Then
                        tdf.Connect = VBA.Replace(tdf.Connect, "D:\Program\Access\market\", Environment.CurrentPath)
                        tdf.RefreshLink
                    End If
                End If
            Next
            Set wsh = Nothing
            Set tdf = Nothing

            AppConfig.Edit
            AppConfig.Fields("Installed") = True
            AppConfig.Update
            AppConfig.Close

            App.Installed = True
            Set AppConfig = Nothing
        End If
    End If

    If Environment.Development = False Then
        ' 隐藏不需要的类别
        DoCmd.SetDisplayedCategories False, ""

        On Error Resume Next
        '导航到指定菜单
        DoCmd.NavigateTo App.AppTitle
        ' 锁定导航窗格
        DoCmd.LockNavigationPane True
    End If

    App.AppStarted = True
    Exit Sub
End Sub


'主函数
Public Function Main() As Boolean
    On Error Resume Next
End Function
