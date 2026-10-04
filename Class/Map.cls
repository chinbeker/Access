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
    Dict.CompareMode = vbBinaryCompare
End Sub

' 检查
Public Function Has(ByVal Key As String) As Boolean
    Key = VBA.Trim(Key)
    If VBA.Len(Key) > 0 Then
        Has = Dict.Exists(Key)
    Else
        Has = False
    End If
End Function

' 读取
Public Property Get Item(ByVal Key As String) As Variant
    Key = VBA.Trim(Key)
    If Has(Key) Then
        Item = Dict.Item(Key)
    Else
        Err.Raise 5, "Map.Item", "键 " & Key & " 不存在"
    End If
End Property

' 赋值
Public Property Let Item(ByVal Key As String, ByVal Value As Variant)
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Property
    Dict.Item(Key) = Value
End Property

' 赋值(对象)
Public Property Set Item(ByVal Key As String, ByVal Value As Object)
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Property
    Set Dict.Item(Key) = Value
End Property

' 计数
Public Property Get Count() As Long
    Count = Dict.Count
End Property

' 添加
Public Sub Add(ByVal Key As String, ByVal Value As Variant)
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Sub
    If Not Has(Key) Then
        Dict.Item(Key) = Value
    Else
        Err.Raise 5, "Map.Add", "键 " & Key & " 已存在"
    End If
End Sub

' 删除
Public Sub Delete(ByVal Key As String)
    Key = VBA.Trim(Key)
    If Has(Key) Then
        Dict.Remove Key
    End If
End Sub

' 清空
Public Sub Clear()
    Dict.RemoveAll
End Sub

' 所有键数组
Public Function Keys() As Variant
    Keys = Dict.Keys
End Function

' 所有值数组
Public Function values() As Variant
    values = Dict.Items
End Function

' 键值对数组
Public Function Entries() As Variant
    Dim result() As Variant
    Dim i As Long
    Dim Key As Variant

    If Dict.count = 0 Then Exit Function
    ReDim result(0 To Dict.count - 1, 0 To 1)
    i = 0
    For Each Key In Dict.Keys
        result(i, 0) = Key
        result(i, 1) = Dict.Item(Key)
        i = i + 1
    Next
    Entries = result
End Function

' 释放
Private Sub Class_Terminate()
    Set Dict = Nothing
End Sub
