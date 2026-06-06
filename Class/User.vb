Option Compare Database
Option Explicit

Public UserName As String

'UserId
Public Property Get UserId() As Long
    If Not GlobalData.Has("UserId") Then GlobalData.SetValue "UserId", 0
    UserId = GlobalData.GetValue("UserId")
End Property
Public Property Let UserId(ByVal value As Long)
    GlobalData.SetValue "UserId", value
End Property
