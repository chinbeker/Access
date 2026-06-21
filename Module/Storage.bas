Attribute VB_Name = "Storage"
'@Lang VBA

Option Compare Database
Option Explicit

Private Function GenerateKey(ByVal key As String) As String
    GenerateKey = "Global_" & key
End Function

'读取
Public Function GetValue(ByVal key As String) As Variant
    If StringBase.IsNullOrEmpty(key) Then GetValue = Null
    GetValue = Application.TempVars.Item(key).value
End Function

'赋值
Public Sub SetValue(ByVal key As String, ByRef value As Variant)
    If Storage.Has(key) Then
        Application.TempVars.Item(key).value = value
    Else
        Application.TempVars.Add key, value
    End If
End Sub

'判断
Public Function Has(ByVal key As String) As Boolean
    If StringBase.IsNullOrEmpty(key) Then Has = False
    Has = Not VBA.IsNull(Application.TempVars.Item(key).value)
End Function

'删除
Public Sub Remove(ByVal key As String)
    If StringBase.IsNullOrEmpty(key) Then Exit Sub
    If Storage.Has(key) Then Application.TempVars.Remove key
End Sub
