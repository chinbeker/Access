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
    Dim count As Integer

    sum = 0
    count = 0

    ' 累加有效数值
    For i = LBound(values) To UBound(values)
        If IsNumeric(values(i)) And Not IsNull(values(i)) Then
            sum = sum + CDbl(values(i))
            count = count + 1
        End If
    Next i

    ' 返回平均值
    If count > 0 Then
        Average = sum / count
    Else
        Average = Null
    End If
End Function

' 生成范围内的随机整数（包含边界）
Public Function Random(ByVal Min As Long, ByVal Max As Long) As Long
    If Max > Min Then
        Random = VBA.Int((Max - Min + 1) * VBA.Rnd + Min)
    Else
        Random = Min
    End If
End Function

' 偶数
Public Function IsEven(ByVal n As Long) As Boolean
    IsEven = (n Mod 2 = 0)
End Function

' 奇数
Public Function IsOdd(ByVal n As Long) As Boolean
    IsOdd = (n Mod 2 <> 0)
End Function
