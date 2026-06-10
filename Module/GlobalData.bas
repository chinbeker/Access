Attribute VB_Name = "GlobalData"
Option Compare Database
Option Explicit

'读取
Public Function GetValue(ByVal key As String) As Variant
    If StringBase.IsWhiteSpace(key) Then GetValue = Null
    GetValue = Application.TempVars.Item(key).value
End Function

'赋值
Public Sub SetValue(ByVal key As String, ByRef value As Variant)
    If StringBase.IsWhiteSpace(key) Then Exit Sub
    Application.TempVars.Add key, value
End Sub

'判断
Public Function Has(ByVal key As String) As Boolean
    If StringBase.IsWhiteSpace(key) Then Has = False
    Has = Not IsNull(Application.TempVars.Item(key).value)
End Function

'删除
Public Sub Remove(ByVal key As String)
    If StringBase.IsWhiteSpace(key) Then Exit Sub
    If Has(key) Then Application.TempVars.Remove key
End Sub
