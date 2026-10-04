Attribute VB_Name = "Utils"
'@Lang VBA

Option Compare Database
Option Explicit


' 多个窗体参数
Public Function CreateOpenArgs(ParamArray values() As Variant) As String
    CreateOpenArgs = ArrayBase.Join(values, "|")
End Function

' 解析窗体参数
Public Function ParseOpenArgs(ByVal str As String) As Variant
    Dim arr As Variant
    If Not StringBase.IsNullOrEmpty(str) Then
        Dim count As Long
        Dim Value As Variant
        Dim i As Long

        arr = VBA.Split(str, "|")
        count = UBound(arr)
        For i = 0 To count
            Value = arr(i)
            If VBA.IsNumeric(Value) Then
                arr(i) = VBA.val(Value)
            ElseIf Value = "True" Or Value = "False" Then
                arr(i) = VBA.CBool(Value)
            ElseIf VBA.Left(Value, 1) = "#" And VBA.Right(Value, 1) = "#" Then
                Value = VBA.Mid(Value, 2, VBA.Len(Value) - 2)
                If VBA.IsDate(Value) Then
                    arr(i) = VBA.CDate(Value)
                Else
                    arr(i) = ""
                End If
            ElseIf Value = "Null" Then
                arr(i) = Null
            End If
        Next i
        ParseOpenArgs = arr
    Else
        ParseOpenArgs = Array()
    End If
End Function
