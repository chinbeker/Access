VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA

Option Compare Database
Option Explicit

'年份
Public Property Get Years() As Long
    If Not Storage.Has("Years") Then
        If Auth.Authorized Then
            Storage.SetValue "Years", DbSql.Lookup("Settings", "Years", "[SysId]=1")
        Else
            Storage.SetValue "Years", 0
        End If
    End If
    Years = Storage.GetValue("Years")
End Property
Public Property Let Years(ByVal value As Long)
    If Not Storage.Has("Years") Or Storage.GetValue("Years") <> value Then
        If DbSql.SetValue("Settings", "Years", value, "[SysId]=1") = True Then Storage.SetValue "Years", value
    End If
End Property

'月份
Public Property Get Months() As Long
    If Not Storage.Has("Months") Then
        If Auth.Authorized Then
            Storage.SetValue "Months", DbSql.Lookup("Settings", "Months", "[SysId]=1")
        Else
            Storage.SetValue "Months", 0
        End If
    End If
    Months = Storage.GetValue("Months")
End Property
Public Property Let Months(ByVal value As Long)
    If Not Storage.Has("Months") Or Storage.GetValue("Months") <> value Then
        If DbSql.SetValue("Settings", "Months", value, "[SysId]=1") = True Then Storage.SetValue "Months", value
    End If
End Property

'工作日ID
Public Property Get WorkShiftId() As Long
    If Not Storage.Has("WorkShiftId") Then
        If Auth.Authorized Then
            Storage.SetValue "WorkShiftId", DbSql.Lookup("Settings", "WorkShiftId", "[SysId]=1")
        Else
            Storage.SetValue "WorkShiftId", 0
        End If
    End If
    WorkShiftId = Storage.GetValue("WorkShiftId")
End Property
Public Property Let WorkShiftId(ByVal value As Long)
    If Not Storage.Has("WorkShiftId") Or Storage.GetValue("WorkShiftId") <> value Then
        If DbSql.SetValue("Settings", "WorkShiftId", value, "[SysId]=1") = True Then Storage.SetValue "WorkShiftId", value
    End If
End Property
