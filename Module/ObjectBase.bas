Attribute VB_Name = "ObjectBase"
'@Lang VBA

Option Compare Database
Option Explicit

' 判断是否是空对象
Public Function IsNothing(ByRef obj As Object) As Boolean
    IsNothing = (obj Is Nothing)
End Function


' 判断一个对象是否为某个类型
Public Function InstanceOf(ByVal obj As Object, ByVal objectType As String) As Boolean
    If obj Is Nothing Then
    Else
        InstanceOf = (TypeName(obj) = objectType)
    End If
End Function


' 判断一个对象是否为数字
Public Function IsNumber(ByVal val As Variant) As Boolean
    IsNumber = IsNumeric(val)
End Function
