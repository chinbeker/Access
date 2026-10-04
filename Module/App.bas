Attribute VB_Name = "App"
'@Lang VBA

Option Compare Database
Option Explicit


'应用名称
Public Function AppName() As String
    AppName = Configuration.GetValue("AppName")
End Function

'应用标题名称
Public Function AppTitle() As String
    AppTitle = Configuration.GetValue("AppTitle")
End Function

'启动状态
Public Function Started() As Boolean
    If State.GetValue("AppStarted") = True Then Started = True
End Function

' 读取配置信息
Public Function GetConfig(ByVal Key As String, Optional ByVal DefaultValue As Variant) As Variant
    GetConfig = Configuration.GetValue(Key, DefaultValue)
End Function

' 添加配置信息
Public Function SetConfig(ByVal Key As String, ByVal Value As Variant) As Boolean
    SetConfig = Configuration.SetValue(Key, Value)
End Function

' 读取状态信息
Public Function GetState(ByVal Key As String, Optional ByVal DefaultValue As Variant) As Variant
    GetState = State.GetValue(Key, DefaultValue)
End Function

' 添加状态信息
Public Function SetState(ByVal Key As String, ByVal Value As Variant) As Boolean
    SetState = State.SetValue(Key, Value)
End Function


' 应用启动
Public Sub Start()
    ' 退出登录
    Auth.Logout

    ' 跳过已启动
    If App.Started Then Exit Sub

    ' 初始化
    If Bootstrap.Initialize Then State.SetValue "AppStarted", True
End Sub
