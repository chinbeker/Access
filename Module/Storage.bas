Attribute VB_Name = "Storage"
'@Lang VBA

Option Compare Database
Option Explicit

'读取
Public Function GetValue(ByVal Key As String, Optional ByVal DefaultValue As Variant) As Variant
    GetValue = Null
    If VBA.Len(Key) = 0 Then Exit Function
    Dim Value As Variant
    Value = Application.TempVars.Item(Key).Value
    If VBA.IsNull(Value) Then
        If Not VBA.IsMissing(DefaultValue) Then GetValue = DefaultValue
    Else
        GetValue = Value
    End If
End Function

'赋值
Public Function SetValue(ByVal Key As String, ByVal Value As Variant) As Boolean
    If VBA.Len(Key) = 0 Then Exit Function
    Application.TempVars(Key) = Value
    SetValue = True
End Function

'判断
Public Function Has(ByVal Key As String) As Boolean
    If VBA.Len(Key) = 0 Then Exit Function
    Has = (Not VBA.IsNull(Application.TempVars.Item(Key).Value))
End Function

'删除
Public Sub Remove(ByVal Key As String)
    If VBA.Len(Key) = 0 Then Exit Sub
    If Storage.Has(Key) Then Application.TempVars.Remove Key
End Sub
