Attribute VB_Name = "State"
'@Lang VBA

Option Compare Database
Option Explicit

'读取
Public Function GetValue(ByVal key As String) As Variant
    If StringBase.IsNullOrEmpty(key) Then GetValue = Null
    GetValue = Application.TempVars.Item("_State_" & key).value
End Function

'赋值
Public Sub SetValue(ByVal key As String, ByRef value As Variant)
    If StringBase.IsNullOrEmpty(key) Then Exit Sub
    Application.TempVars.Add "_State_" & key, value
End Sub

'判断
Public Function Has(ByVal key As String) As Boolean
    If StringBase.IsNullOrEmpty(key) Then Has = False
    Has = Not IsNull(Application.TempVars.Item("_State_" & key).value)
End Function

'删除
Public Sub Remove(ByVal key As String)
    If StringBase.IsNullOrEmpty(key) Then Exit Sub
    If Has("_State_" & key) Then Application.TempVars.Remove "_State_" & key
End Sub
