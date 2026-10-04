VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlainObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA

Option Compare Database
Option Explicit

Private Data As New Collection

' 检查
Public Function Has(ByVal Key As String) As Boolean
    Dim val As Variant
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Function
    Err.Clear
    On Error Resume Next
    val = Data.Item(Key)
    Has = (Err.Number = 0)
    On Error GoTo 0
End Function

' 读取
Public Property Get Prop(ByVal Key As String) As Variant
    Key = VBA.Trim(Key)
    If Has(Key) Then
        Prop = Data.Item(Key)
    Else
        Err.Raise vbObjectError + 1001, "HashSet.Prop", "键 " & Key & " 不存在"
    End If
End Property

' 赋值
Public Property Let Prop(ByVal Key As String, ByVal Value As Variant)
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Property
    If Has(Key) Then Data.Remove Key
    Data.Add Value, Key
End Property

' 赋值(对象)
Public Property Set Prop(ByVal Key As String, ByVal Value As Object)
    Key = VBA.Trim(Key)
    If VBA.Len(Key) = 0 Then Exit Property
    If Has(Key) Then Data.Remove Key
    Data.Add Value, Key
End Property

' 计数
Public Property Get Count() As Long
    Count = Data.Count
End Property

' 删除
Public Sub Remove(ByVal Key As String)
    Key = VBA.Trim(Key)
    If Has(Key) Then Data.Remove Key
End Sub

' 清空
Public Sub Clear()
    Set Data = New Collection
End Sub

' 释放
Private Sub Class_Terminate()
    Set Data = Nothing
End Sub
