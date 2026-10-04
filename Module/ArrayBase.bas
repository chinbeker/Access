Attribute VB_Name = "ArrayBase"
'@Lang VBA

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
Public Function Length(ByRef arr As Variant) As Long
    If Not IsArray(arr) Then
        Length = -1
    Else
        On Error Resume Next
        Length = UBound(arr) - LBound(arr) + 1
    End If
End Function


' 数组拼接成字符串
Public Function Join(ByVal arr As Variant, Optional ByVal separator As String) As String
    If VBA.IsMissing(separator) Or VBA.Len(separator) = 0 Then separator = ","
    If Not ArrayBase.IsEmpty(arr) Then
        Dim text As String
        Dim Value As String
        Dim i As Long
        Dim count As Long

        count = UBound(arr) - LBound(arr)
        If count > 0 Then
            For i = LBound(arr) To UBound(arr)

                Select Case VBA.VarType(arr(i))
                    Case VBA.vbString
                        Value = arr(i)
                    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal, VBA.vbByte, VBA.vbBoolean
                        Value = VBA.CStr(arr(i))
                    Case VBA.vbDate
                        Value = "#" & VBA.Format(arr(i), "yyyy-mm-dd") & "#"
                End Select

                If i < count Then
                    text = text & Value & separator
                Else
                    text = text & Value
                End If
            Next i
            Join = text
        End If
    End If
End Function
