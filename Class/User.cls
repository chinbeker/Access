VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA


Option Compare Database
Option Explicit

Public UserName As String

'UserId
Public Property Get UserId() As Long
    If Not Storage.Has("UserId") Then Storage.SetValue "UserId", 0
    UserId = Storage.GetValue("UserId")
End Property
Public Property Let UserId(ByVal value As Long)
    Storage.SetValue "UserId", value
End Property
