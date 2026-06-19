Attribute VB_Name = "StringBase"
'@Lang VBA


Option Compare Database
Option Explicit

' 判断一个变量是否为字符串类型
Public Function IsString(ByVal str As Variant) As Boolean
    IsString = (VBA.VarType(str) = VBA.vbString)
End Function

' 判断一个字符串是否为空字符串
Public Function IsNullOrEmpty(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsNullOrEmpty = (VBA.Len(str) = 0)
    Else
        IsNullOrEmpty = True
    End If
End Function

' 字符串比较（区分大小写）
Public Function Equals(ByVal str1 As String, ByVal str2 As String) As Boolean
    Equals = (VBA.StrComp(str1, str2, vbBinaryCompare) = 0)
End Function

' 字符串长度
Public Function Length(ByVal str As Variant) As Long
    If Not IsNullOrEmpty(str) Then
        Length = VBA.Len(str)
    Else
        Length = 0
    End If
End Function

' 判断字符串是否包含子串
Public Function Contains(ByVal str As Variant, ByVal search As String) As Boolean
    If Not IsNullOrEmpty(str) Then
        Contains = (VBA.InStr(1, str, search, vbBinaryCompare) > 0)
    Else
        Contains = False
    End If
End Function

' 查找子串在字符串中第一次出现的位置
Public Function IndexOf(ByVal str As Variant, ByVal search As String) As Long
    If Not IsNullOrEmpty(str) Then
        Dim pos As Long
        pos = VBA.InStr(1, str, search, vbBinaryCompare)
        If pos > 0 Then
            IndexOf = pos - 1
        Else
            IndexOf = -1
        End If
    Else
        IndexOf = -1
    End If
End Function

' 查找子串在字符串末尾中第一次出现的位置
Public Function LastIndexOf(ByVal str As Variant, ByVal search As String) As Long
    If Not IsNullOrEmpty(str) Then
        Dim pos As Long
        pos = VBA.InStrRev(str, search, -1, vbBinaryCompare)
        If pos > 0 Then
            LastIndexOf = pos - 1
        Else
            LastIndexOf = -1
        End If
    Else
        LastIndexOf = -1
    End If
End Function


'提取字符串
Public Function Substring(ByVal str As Variant, ByVal start As Long, Optional ByVal Length As Long = -1) As String
    If Not IsNullOrEmpty(str) Then
        Dim strLen As Long
        strLen = VBA.Len(str)
        If start < 0 Then start = 0
        If start > strLen Then start = strLen
        If Length < 0 Then
            Substring = VBA.Mid(str, start + 1)
        Else
            If start + Length <= strLen Then
                 Substring = VBA.Mid(str, start + 1, Length)
            Else
                 Substring = VBA.Mid(str, start + 1)
            End If
        End If
    End If
End Function

' 判断字符串是否以指定子串开头（区分大小写）
Public Function StartsWith(ByVal str As Variant, ByVal value As String) As Boolean
    If Not IsNullOrEmpty(str) Then
        Dim ValueLength As Long
        ValueLength = VBA.Len(value)
        If ValueLength > 0 Then
            StartsWith = Equals(VBA.Left(str, ValueLength), value)
        Else
            StartsWith = True
        End If
    End If
End Function


' 判断字符串是否以指定子串结尾（区分大小写）
Public Function EndsWith(ByVal str As Variant, ByVal value As String) As Boolean
    If Not IsNullOrEmpty(str) Then
        Dim ValueLength As Long
        ValueLength = VBA.Len(value)
        If ValueLength > 0 Then
            EndsWith = Equals(VBA.Right(str, ValueLength), value)
        Else
            EndsWith = True
        End If
    End If
End Function


' 字符串替换
Public Function Replace(ByVal str As Variant, ByVal oldValue As String, ByVal newValue As String) As String
    If Not IsNullOrEmpty(str) Then
        Replace = VBA.Replace(str, oldValue, newValue, 1, -1, vbBinaryCompare)
    End If
End Function

' 转换为大写
Public Function ToUpper(ByVal str As Variant) As String
    If Not IsNullOrEmpty(str) Then ToUpper = VBA.UCase(str)
End Function

' 转换为小写
Public Function ToLower(ByVal str As Variant) As String
    If Not IsNullOrEmpty(str) Then ToLower = VBA.LCase(str)
End Function


' 去除字符串开头的空白字符
Public Function TrimStart(ByVal str As Variant) As String
    If Not IsNullOrEmpty(str) Then
        Dim result As String
        Dim char As String
        result = str
        Do While VBA.Len(result) > 0
            char = VBA.Left(result, 1)
            If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
                result = VBA.Mid(result, 2)
            Else
                Exit Do
            End If
        Loop
        TrimStart = result
    End If
End Function

' 去除字符串结尾的空白字符
Public Function TrimEnd(ByVal str As Variant) As String
    If Not IsNullOrEmpty(str) Then
        Dim result As String
        Dim char As String
        result = str
        Do While VBA.Len(result) > 0
            char = VBA.Right(result, 1)
            If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
                result = VBA.Left(result, VBA.Len(result) - 1)
            Else
                Exit Do
            End If
        Loop
        TrimEnd = result
    End If
End Function

' 去除字符串首尾的空白字符
Public Function Trim(ByVal str As Variant) As String
    If Not IsNullOrEmpty(str) Then Trim = TrimEnd(TrimStart(str))
End Function


' 判断一个字符串是否为空白字符串
Public Function IsNullOrWhiteSpace(ByVal str As Variant) As Boolean
    If Not IsNullOrEmpty(str) Then
        IsNullOrWhiteSpace = (VBA.Len(StringBase.Trim(str)) = 0)
    Else
        IsNullOrWhiteSpace = True
    End If
End Function

' 判断一个字符串是否为空字符串，并去除首位空白字符
Public Function IsNullOrWhiteSpaceAndTrim(ByRef str As Variant) As Boolean
    If Not IsNullOrEmpty(str) Then
        str = StringBase.Trim(str)
        IsNullOrWhiteSpaceAndTrim = (VBA.Len(str) = 0)
    Else
        IsNullOrWhiteSpaceAndTrim = True
    End If
End Function
