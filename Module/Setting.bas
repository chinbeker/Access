Attribute VB_Name = "Setting"
'@Lang VBA

Option Compare Database
Option Explicit

Private Const TABLE_NAME As String = "Settings"
Private Const PREFIX As String = "Setting_"

'检查
Public Function Has(ByVal Key As String) As Boolean
    If VBA.Len(Key) = 0 Then Exit Function
    Dim SettingKey As String
    SettingKey = PREFIX & Key
    Has = (Not VBA.IsNull(Application.TempVars.Item(Key).Value))
End Function

'读取
Public Function GetValue(ByVal Key As String, Optional ByVal DefaultValue As Variant) As Variant
    GetValue = Null
    If VBA.Len(Key) = 0 Then Exit Function

    Dim SettingKey As String
    SettingKey = PREFIX & Key

    Dim Value As Variant
    Value = Application.TempVars.Item(SettingKey).Value

    If VBA.IsNull(Value) Then
        If Auth.Authorized Then
            Storage.SetValue SettingKey, DbSql.Lookup(TABLE_NAME, Key, "[SysId]=1")
        ElseIf VBA.IsMissing(DefaultValue) Then
            Storage.SetValue SettingKey, Null
        Else
            Storage.SetValue SettingKey, DefaultValue
        End If
        GetValue = Storage.GetValue(SettingKey)
    Else
        GetValue = Value
    End If
End Function

'赋值
Public Function SetValue(ByVal Key As String, ByVal Value As Variant) As Boolean
    If VBA.Len(Key) = 0 Then Exit Function
    Dim SettingKey As String
    SettingKey = PREFIX & Key
    If Storage.Has(SettingKey) Then
        If Storage.GetValue(SettingKey) <> Value Then
             If DbSql.SetValue(TABLE_NAME, Key, Value, "[SysId]=1") = True Then Storage.SetValue SettingKey, Value
        End If
    Else
        Storage.SetValue SettingKey, Value
    End If
End Function



'----------------------自定义快速访问-------------------------
' Years
Public Function Years() As Long
    Years = Setting.GetValue("Years", 0)
End Function

' LunarYear
Public Function LunarYear() As Long
    LunarYear = Setting.GetValue("LunarYear", 0)
End Function

' Months
Public Function Months() As Long
    Months = Setting.GetValue("Months", 0)
End Function
