Attribute VB_Name = "Bootstrap"
'@Lang VBA

Option Compare Database
Option Explicit


'已安装
Public Function AppInstalled() As Boolean
    If Configuration.GetValue("AppInstalled") = True Then AppInstalled = True
End Function

'启用SqlServer
Public Function UseSqlServer() As Boolean
    If Configuration.GetValue("UseSqlServer") = True Then UseSqlServer = True
End Function


'安装 ODBC_Driver_18_for_SQL_Server 驱动
Private Sub InstallODBCDriver()
    On Error GoTo ErrorHandler
    Dim wsh As Object
    Dim path As String
    Dim val As Variant
    Dim installed As Boolean

    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    val = wsh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\ODBC Driver 18 for SQL Server")
    installed = VBA.IIf(Err.Number = 0, True, False)

    On Error GoTo ErrorHandler
    If Not installed Then
        path = Environment.CurrentPath & "resource\driver\Microsoft_ODBC_Driver_18_for_SQL_Server_18.6.2_x64.msi"
        If VBA.Dir(path) = "" Then
            Message.Warning "未找到 Microsoft_ODBC_Driver_18_for_SQL_Server 驱动文件"
            Exit Sub
        End If
        wsh.Run "msiexec.exe /i """ & path & """ /norestart", 1, True
    End If

    Set wsh = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

'安装 OLE_DB_Driver_18_for_SQL_Server 驱动
Private Sub InstallOLEDBDriver()
    On Error GoTo ErrorHandler
    Dim wsh As Object
    Dim path As String
    Dim val As Variant
    Dim installed As Boolean

    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    val = wsh.RegRead("HKLM\SOFTWARE\Microsoft\MSOLEDBSQL\InstalledVersion")
    installed = VBA.IIf(Err.Number = 0, True, False)

    On Error GoTo ErrorHandler
    If Not installed Or val <> "18.7.5.0" Then
        path = Environment.CurrentPath & "resource\driver\Microsoft_OLE_DB_Driver_18_for_SQL_Server_18.7.5_x64.msi"
        If VBA.Dir(path) = "" Then
            Message.Warning "未找到 Microsoft_OLE_DB_Driver_18_for_SQL_Server 驱动文件"
            Exit Sub
        End If
        wsh.Run "msiexec.exe /i """ & path & """ /norestart", 1, True
    End If

    Set wsh = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

'导入 SQLServer ODBC 数据源（DSN）
Private Sub InstallDataSource()
    On Error GoTo ErrorHandler
    Dim wsh As Object
    Dim path As String

    Set wsh = CreateObject("WScript.Shell")

    ' 数据源文件目录
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

    Set wsh = Nothing
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub


'刷新失效的链接表
Private Sub RefreshLink()
    On Error Resume Next

    Dim DB As DAO.Database
    Dim tdf As DAO.TableDef
    Dim count As Long
    Dim size As Integer
    Dim val As Integer

    '链接表计数
    Progress.Start "系统启动中"
    Set DB = Application.CurrentDb
    Progress.Run 5
    For Each tdf In DB.TableDefs
        If Not (VBA.Left(tdf.name, 4) = "MSys" Or VBA.Left(tdf.name, 4) = "USys") Then
            If VBA.Len(tdf.Connect) > 0 Then
                If (tdf.Attributes And dbAttachedTable) <> 0 And (tdf.Attributes And dbAttachedODBC) = 0 Then
                    count = count + 1
                End If
            End If
        End If
    Next tdf
    Progress.Run 10

    '刷新链接
    If count > 0 Then
        size = VBA.CInt(80 / count)
        val = 15
        For Each tdf In DB.TableDefs
            If Not (VBA.Left(tdf.name, 4) = "MSys" Or VBA.Left(tdf.name, 4) = "USys") Then
                If VBA.Len(tdf.Connect) > 0 Then
                    If (tdf.Attributes And dbAttachedTable) <> 0 And (tdf.Attributes And dbAttachedODBC) = 0 Then
                        tdf.Connect = VBA.Replace(tdf.Connect, "D:\Program\Access\", Environment.ParentPath)
                        tdf.RefreshLink
                        val = val + size
                        Progress.Run val
                    End If
                End If
            End If
        Next tdf
    Else
        Progress.Run 90
    End If

    Set DB = Nothing
    Set tdf = Nothing
End Sub


'安装驱动
Private Sub Install()
    On Error GoTo ErrorHandler

    If Environment.Development Then Exit Sub                '跳过开发环境
    If Bootstrap.AppInstalled Then Exit Sub                 '跳过已安装

    '使用 SqlServer 进入安装
    If UseSqlServer Then
        Call InstallODBCDriver                              '安装 ODBC 驱动
        Call InstallOLEDBDriver                             '安装 OLEDB 驱动
        'Call InstallDataSource                             '安装 DSN 数据源

        On Error GoTo ErrorHandler
        Dim AppConfig As DAO.Recordset
        Set AppConfig = DbSql.TableFirstRecord("Application")
        AppConfig.Edit
        AppConfig.Fields("Installed").Value = True
        Configuration.SetValue "AppInstalled", True         '配置为已安装
        AppConfig.Update
        AppConfig.Close
        Set AppConfig = Nothing
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub


'初始化
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler
    Dim AppConfig As DAO.Recordset

    ' 禁用执行 Sql 时的弹窗提示
    DoCmd.SetWarnings (False)

    '读取应用配置
    Set AppConfig = DbSql.TableFirst("Application")
    If Not AppConfig Is Nothing Then
        If AppConfig.RecordCount > 0 Then
            Configuration.SetValue "AppName", AppConfig.Fields("AppName").Value
            Configuration.SetValue "AppTitle", AppConfig.Fields("AppTitle").Value
            Configuration.SetValue "CurrentPath", AppConfig.Fields("CurrentPath").Value
            Configuration.SetValue "Installed", AppConfig.Fields("Installed").Value
            Configuration.SetValue "Development", AppConfig.Fields("Development").Value
            Configuration.SetValue "UseSqlServer", AppConfig.Fields("UseSqlServer").Value
        End If
    Else
        Message.Warning "未找到应用程序基本配置信息"
    End If
    AppConfig.Close
    Set AppConfig = Nothing

    '锁定导航窗格
    On Error Resume Next
    If Not Environment.Development Then
        DoCmd.SetDisplayedCategories False, ""                             ' 隐藏不需要的类别
        DoCmd.NavigateTo "导航菜单"                                        ' 导航到指定菜单
        DoCmd.LockNavigationPane True
    End If

    '检查安装目录
    On Error GoTo ErrorHandler
    If Configuration.GetValue("CurrentPath") <> Environment.CurrentPath Then
        '刷新链接表
        Call RefreshLink

        '更新配置
        Configuration.SetValue "CurrentPath", Environment.CurrentPath
        Set AppConfig = DbSql.TableFirstRecord("Application")
        If Not AppConfig Is Nothing Then
            If AppConfig.RecordCount > 0 Then
                AppConfig.Edit
                AppConfig.Fields("CurrentPath").Value = Environment.CurrentPath
                AppConfig.Update
            End If
        End If
        AppConfig.Close
        Set AppConfig = Nothing

        Progress.Run 100
        Progress.Complete
    End If


    '安装 SqlServer 运行环境
    Call Install

    Initialize = True
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Initialize = False
End Function
