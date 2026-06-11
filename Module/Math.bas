Attribute VB_Name = "Math"
'@Lang VBA

Option Compare Database
Option Explicit

' 最小值
Function Min(ParamArray values() As Variant) As Variant
    Dim i As Integer
    Dim minVal As Variant

    minVal = values(LBound(values))

    ' 循环比较
    For i = LBound(values) + 1 To UBound(values)
        If values(i) < minVal Then
            minVal = values(i)
        End If
    Next i

    Min = minVal
End Function

' 最大值
Function Max(ParamArray values() As Variant) As Variant
    Dim i As Integer
    Dim maxVal As Variant

    maxVal = values(LBound(values))

    For i = LBound(values) + 1 To UBound(values)
        If values(i) > maxVal Then
            maxVal = values(i)
        End If
    Next i

    Max = maxVal
End Function

' 平均值
Function Average(ParamArray values() As Variant) As Variant
    Dim i As Integer
    Dim sum As Double
    Dim Count As Integer

    sum = 0
    Count = 0

    ' 累加有效数值
    For i = LBound(values) To UBound(values)
        If IsNumeric(values(i)) And Not IsNull(values(i)) Then
            sum = sum + CDbl(values(i))
            Count = Count + 1
        End If
    Next i

    ' 返回平均值
    If Count > 0 Then
        Average = sum / Count
    Else
        Average = Null
    End If
End Function
