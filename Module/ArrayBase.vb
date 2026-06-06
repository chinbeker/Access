Option Compare Database
Option Explicit

' 判断一个数组是否为空数组
Public Function IsEmpty(ByRef arr As Variant) As Boolean
    If IsArray(arr) Then
        IsEmpty = (LBound(arr) > UBound(arr))
    Else
        IsEmpty = False
    End If
End Function


' 返回数组长度
Function length(ByRef arr As Variant) As Long
    If Not IsArray(arr) Then
        length = -1
    Else
        On Error Resume Next
        length = UBound(arr) - LBound(arr) + 1
    End If
End Function

' 字符串拼接
Function Join(ByRef arr As Variant, Optional ByVal separator As String = ",") As String
    Dim text As String
    If Not IsEmpty(arr) Then
        Dim i As Long
        Dim length As Long

        length = UBound(arr) - LBound(arr) + 1

        For i = LBound(arr) To UBound(arr)
            If i < length Then
                text = text & arr(i) & separator
            Else
                text = text & arr(i)
            End If
        Next i
    End If
    Join = text
End Function
