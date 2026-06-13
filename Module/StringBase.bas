Attribute VB_Name = "StringBase"
'@Lang VBA

Option Compare Database
Option Explicit

' 判断一个变量是否为字符串类型
Public Function IsString(ByVal str As Variant) As Boolean
    On Error GoTo ErrorHandler
    IsString = (VBA.VarType(str) = VBA.vbString)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 判断一个字符串是否为空字符串
Public Function IsWhiteSpace(ByVal str As Variant) As Boolean
    On Error GoTo ErrorHandler
    If VBA.VarType(str) = VBA.vbString Then
        IsWhiteSpace = (VBA.Len(VBA.Trim(str)) = 0)
    Else
        IsWhiteSpace = True
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 字符串比较（区分大小写）
Public Function Compare(ByVal str1 As String, ByVal str2 As String) As Boolean
    On Error GoTo ErrorHandler
    Compare = (VBA.StrComp(str1, str2, vbBinaryCompare) = 0)
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
