VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA

Option Compare Database
Option Explicit

Private Dict As Object

Private Sub Class_Initialize()
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = 0
End Sub

' 读取
Public Property Get Item(ByVal key As String) As Variant
    key = VBA.Replace(key, " ", "")
    If Has(key) Then
        Item = Dict.Item(key)
    Else
        Err.Raise 5, "Map.Item", "键 " & key & " 不存在"
    End If
End Property

' 赋值
Public Property Let Item(ByVal key As String, ByVal value As Variant)
    key = VBA.Replace(key, " ", "")
    If VBA.Len(key) = 0 Then Exit Property
    Dict.Item(key) = value
End Property

' 赋值(对象)
Public Property Set Item(ByVal key As String, ByVal value As Object)
    key = VBA.Replace(key, " ", "")
    If VBA.Len(key) = 0 Then Exit Property
    Set Dict.Item(key) = value
End Property

' 添加
Public Sub Add(ByVal key As String, ByVal value As Variant)
    key = VBA.Replace(key, " ", "")
    If VBA.Len(key) = 0 Then Exit Sub
    Dict.Item(key) = value
End Sub

' 删除
Public Sub Delete(ByVal key As String)
    key = VBA.Replace(key, " ", "")
    If Has(key) Then
        Dict.Remove key
    End If
End Sub

' 计数
Public Function Count() As Long
    Count = Dict.Count
End Function

' 清空
Public Sub Clear()
    Dict.RemoveAll
End Sub

' 检查
Public Function Has(ByVal key As String) As Boolean
    key = VBA.Replace(key, " ", "")
    If VBA.Len(key) > 0 Then
        Has = Dict.Exists(key)
    Else
        Has = False
    End If
End Function

' 所有键数组
Public Function Keys() As Variant
    Keys = Dict.Keys
End Function

' 所有值数组
Public Function Values() As Variant
    Values = Dict.Items
End Function

' 键值对数组
Public Function Entries() As Variant
    Dim result() As Variant
    Dim i As Long
    Dim key As Variant

    ReDim result(0 To Dict.Count - 1, 0 To 1)
    i = 0
    For Each key In Dict.Keys
        result(i, 0) = key
        result(i, 1) = Dict.Item(key)
        i = i + 1
    Next
    Entries = result
End Function

' 释放
Private Sub Class_Terminate()
    Dict.RemoveAll
    Set Dict = Nothing
End Sub
