Attribute VB_Name = "State"
'@Lang VBA

Option Compare Database
Option Explicit


Private Function GenerateKey(ByVal key As String) As String
    GenerateKey = "State_" & key
End Function


'读取
Public Function GetValue(ByVal key As String) As Variant
    If StringBase.IsNullOrEmpty(key) Then GetValue = Null
    GetValue = Application.TempVars.Item(GenerateKey(key)).value
End Function

'赋值
Public Sub SetValue(ByVal key As String, ByRef value As Variant)
    If State.Has(key) Then
        Application.TempVars.Item(GenerateKey(key)).value = value
    Else
        Application.TempVars.Add GenerateKey(key), value
    End If
End Sub

'判断
Public Function Has(ByVal key As String) As Boolean
    If StringBase.IsNullOrEmpty(key) Then Has = False
    Has = Not VBA.IsNull(Application.TempVars.Item(GenerateKey(key)).value)
End Function

'删除
Public Sub Remove(ByVal key As String)
    If StringBase.IsNullOrEmpty(key) Then Exit Sub
    key = GenerateKey(key)
    If State.Has(key) Then Application.TempVars.Remove key
End Sub
