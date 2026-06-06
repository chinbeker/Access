Option Compare Database
Option Explicit

'年份
Public Property Get Years() As Long
    If Not GlobalData.Has("Years") Then
        If App.DatabaseConnected Then
            GlobalData.SetValue "Years", DbSql.Lookup("Settings", "Years", "[SysId]=1")
        Else
            GlobalData.SetValue "Years", 0
        End If
    End If
    Years = GlobalData.GetValue("Years")
End Property
Public Property Let Years(ByVal value As Long)
    If Not GlobalData.Has("Years") Or GlobalData.GetValue("Years") <> value Then
        If DbSql.SetValue("Settings", "Years", value, "[SysId]=1") = True Then GlobalData.SetValue "Years", value
    End If
End Property

'月份
Public Property Get Months() As Long
    If Not GlobalData.Has("Months") Then
        If App.DatabaseConnected Then
            GlobalData.SetValue "Months", DbSql.Lookup("Settings", "Months", "[SysId]=1")
        Else
            GlobalData.SetValue "Months", 0
        End If
    End If
    Months = GlobalData.GetValue("Months")
End Property
Public Property Let Months(ByVal value As Long)
    If Not GlobalData.Has("Months") Or GlobalData.GetValue("Months") <> value Then
        If DbSql.SetValue("Settings", "Months", value, "[SysId]=1") = True Then GlobalData.SetValue "Months", value
    End If
End Property

'工作日ID
Public Property Get WorkShiftId() As Long
    If Not GlobalData.Has("WorkShiftId") Then
        If App.DatabaseConnected Then
            GlobalData.SetValue "WorkShiftId", DbSql.Lookup("Settings", "WorkShiftId", "[SysId]=1")
        Else
            GlobalData.SetValue "WorkShiftId", 0
        End If
    End If
    WorkShiftId = GlobalData.GetValue("WorkShiftId")
End Property
Public Property Let WorkShiftId(ByVal value As Long)
    If Not GlobalData.Has("WorkShiftId") Or GlobalData.GetValue("WorkShiftId") <> value Then
        If DbSql.SetValue("Settings", "WorkShiftId", value, "[SysId]=1") = True Then GlobalData.SetValue "WorkShiftId", value
    End If
End Property
