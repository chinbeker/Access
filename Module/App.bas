Attribute VB_Name = "App"
'@Lang VBA

Option Compare Database
Option Explicit

Public AppName As String             '应用名称
Public AppTitle As String            '标题名称
Public Connected As Boolean          '数据库连接状态
Public Settings As Setting           '设置

'已启动
Public Function Started() As Boolean
    If State.GetValue("AppStarted") = True Then Started = True
End Function
'已安装
Public Function Installed() As Boolean
    If State.GetValue("AppInstalled") = True Then Installed = True
End Function

'初始化
Public Sub Initialize()
    Dim AppConfig As DAO.Recordset
    Set AppConfig = DbSql.TableFirst("Application")

    If Not AppConfig Is Nothing And Not AppConfig.EOF Then
        App.AppName = AppConfig.Fields("AppName").value
        App.AppTitle = AppConfig.Fields("AppTitle").value
        State.SetValue "AppInstalled", AppConfig.Fields("Installed").value
        State.SetValue "Development", AppConfig.Fields("Development").value
    Else
        Message.Warning "未找到应用程序基本配置信息"
    End If

    If App.Settings Is Nothing Then Set App.Settings = New Setting
    AppConfig.Close
    Set AppConfig = Nothing
End Sub


' 应用启动
Public Sub Start()
    '禁用执行 Sql 时的弹窗提示
    DoCmd.SetWarnings (False)

    ' 清除用户登录状态
    State.SetValue "Authenticated", False

    '初始化
    App.Initialize

    '跳过已启动
    If App.Started Then Exit Sub

    '首次安装
    App.Install

    '生产环境执行代码
    If Not Environment.Development Then
        DoCmd.SetDisplayedCategories False, ""      ' 隐藏不需要的类别
        On Error Resume Next
        DoCmd.NavigateTo App.AppTitle               ' 导航到指定菜单
        DoCmd.LockNavigationPane True               ' 锁定导航窗格
    End If
    State.SetValue "AppStarted", True               ' 标记为已启动
End Sub

'首次安装
Private Sub Install()
    If Environment.Development Then Exit Sub
    If App.Installed Then Exit Sub

    Dim wsh As Object
    Dim path As String
    Dim value As String
    Dim IsDriverInstalled As Boolean

    '信任安全位置
    Set wsh = CreateObject("WScript.Shell")
    path = "HKCU\SOFTWARE\Microsoft\Office\16.0\Access\Security\Trusted Locations\Location8\"

    On Error Resume Next
    value = wsh.RegRead(path)
    IsDriverInstalled = VBA.IIf(err.Number = 0, True, False)

    On Error GoTo 0
    '如果已经注册，跳过注册
    If IsDriverInstalled Then Exit Sub

    '写入注册表
    wsh.RegWrite path & "Path", Environment.CurrentPath, "REG_SZ"
    wsh.RegWrite path & "AllowSubfolders", 1, "REG_DWORD"
    wsh.RegWrite path & "Description", "", "REG_SZ"
    wsh.RegWrite path & "Date", "02/01/2026 12:00", "REG_SZ"

    '安装 ODBC_Driver_18_for_SQL_Server 驱动
    value = ""
    IsDriverInstalled = False

    On Error Resume Next
    value = wsh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\ODBC Driver 18 for SQL Server")
    IsDriverInstalled = VBA.IIf(err.Number = 0, True, False)

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
    IsDriverInstalled = VBA.IIf(err.Number = 0, True, False)

    On Error GoTo 0
    If Not IsDriverInstalled Or value <> "18.7.5.0" Then
        path = Environment.CurrentPath & "resource\driver\Microsoft_OLE_DB_Driver_18_for_SQL_Server_18.7.5_x64.msi"
        If VBA.Dir(path) = "" Then
            Message.Warning "未找到 Microsoft_OLE_DB_Driver_18_for_SQL_Server 驱动文件"
            Exit Sub
        End If
        wsh.Run "msiexec.exe /i """ & path & """ /passive /norestart", 0, True
    End If

    '导入 ODBC 数据源配置
    path = Environment.CurrentPath & "config\ODBC_Data_Sources.reg"
    If VBA.Dir(path) = "" Then
        Message.Warning "未找到 ODBC 数据源配置文件"
        Exit Sub
    End If

    On Error Resume Next
    wsh.Run "regedit.exe /s """ & path & """", 0, True
    If err.Number <> 0 Then
        Message.Warning "配置文件导入失败，请联系管理员"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
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

    ' 保存安装记录
    Dim AppConfig As DAO.Recordset
    Set AppConfig = DbSql.TableFirstRecord("Application")

    AppConfig.Edit
    AppConfig.Fields("Installed").value = True
    State.SetValue "AppInstalled", True
    AppConfig.Update
    AppConfig.Close

    Set wsh = Nothing
    Set tdf = Nothing
    Set AppConfig = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(err)
    Exit Sub
End Sub

'主函数
Public Function Main() As Boolean
    On Error Resume Next
End Function
