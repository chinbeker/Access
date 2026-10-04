Attribute VB_Name = "Security"
'@Lang VBA

Option Compare Database
Option Explicit

Public Function SHA256(ByVal PlainText As String) As String
    Dim encoder As Object
    Dim hasher As Object
    Dim textToHash() As Byte
    Dim hash() As Byte
    Dim cypher() As String
    Dim x As Long, l As Long, u As Long

    On Error GoTo CleanUp
    Set encoder = CreateObject("System.Text.UTF8Encoding")
    Set hasher = CreateObject("System.Security.Cryptography.SHA256Managed")
    textToHash = encoder.GetBytes_4(PlainText)
    hash = hasher.ComputeHash_2(textToHash)
    l = LBound(hash)
    u = UBound(hash)
    ReDim cypher(l To u)
    For x = l To u
        cypher(x) = VBA.Right("0" & VBA.Hex$(VBA.CLng(hash(x))), 2)
    Next x
    Set hasher = Nothing
    Set encoder = Nothing
    SHA256 = VBA.Join(cypher, "")
    Exit Function

CleanUp:
    Set hasher = Nothing
    Set encoder = Nothing
    Message.Warning "安全模块 SHA256 无效"
    Exit Function
End Function
