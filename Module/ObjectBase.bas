Attribute VB_Name = "ObjectBase"
'@Lang VBA

Option Compare Database
Option Explicit

' 判断是否是空对象
Public Function IsNothing(ByRef obj As Object) As Boolean
    IsNothing = (obj Is Nothing)
End Function


' 判断一个对象是否为某个类型
Public Function InstanceOf(ByRef obj As Object, ByVal objectType As String) As Boolean
    If obj Is Nothing Then
        InstanceOf = False
    Else
        InstanceOf = (VBA.TypeName(obj) = objectType)
    End If
End Function


' 判断一个对象是否为数字
Public Function IsNumber(ByVal val As Variant) As Boolean
    IsNumber = VBA.IsNumeric(val)
End Function

' 判断某个对象是否 Dictionary 字典
Public Function IsDictionary(ByRef obj As Variant) As Boolean
    IsDictionary = VBA.IsObject(obj) And VBA.TypeName(obj) = "Dictionary"
End Function

' 判断某个对象是否 Collection 集合
Public Function IsCollection(ByRef obj As Variant) As Boolean
    IsCollection = VBA.IsObject(obj) And VBA.TypeName(obj) = "Collection"
End Function
