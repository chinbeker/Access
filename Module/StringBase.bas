Attribute VB_Name = "StringBase"
'@Lang VBA

Option Compare Database
Option Explicit

' 判断一个变量是否为字符串类型
Public Function IsString(ByVal str As Variant) As Boolean
    IsString = (VBA.VarType(str) = VBA.vbString)
End Function

' 判断一个字符串是否为空字符串
Public Function IsEmpty(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsEmpty = (VBA.Len(str) = 0)
    Else
        IsEmpty = True
    End If
End Function

' 判断一个字符串是否为空字符串
Public Function IsWhiteSpace(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsWhiteSpace = (VBA.Len(VBA.Trim(str)) = 0)
    Else
        IsWhiteSpace = True
    End If
End Function

' 字符串比较（区分大小写）
Public Function Compare(ByVal str1 As String, ByVal str2 As String) As Boolean
    Compare = (VBA.StrComp(str1, str2, vbBinaryCompare) = 0)
End Function
