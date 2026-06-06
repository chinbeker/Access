Option Compare Database
Option Explicit

Public Development As Boolean        '开发环境

'生产环境
Public Function Production() As Boolean
    Production = Not Development
End Function

'系统桌面路径
Public Function DesktopPath() As String
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

'当前安装路径
Public Function CurrentPath() As String
    CurrentPath = Application.CurrentProject.path & "\"
End Function

'计算机名称
Public Function ComputerName() As String
    ComputerName = VBA.Environ("COMPUTERNAME")
End Function

'计算机用户名
Public Function UserName() As String
    UserName = VBA.Environ("USERNAME")
End Function

'处理器架构
Public Function ProcessorArchitecture() As String
    ProcessorArchitecture = VBA.Environ("PROCESSOR_ARCHITECTURE")
End Function

'系统内核
Public Function OS() As String
    OS = VBA.Environ("OS")
End Function
