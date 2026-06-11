Attribute VB_Name = "HTTP"
'@Lang VBA

Option Compare Database
Option Explicit

Function Request(ByVal url As String) As String
    On Error GoTo ErrorHandler
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' 异步设置为 False，表示同步请求，代码会等待响应
    req.Open "GET", url, False
    req.SetTimeouts 3000
    req.Send

    ' 检查状态码
    If req.Status = 200 Then
        Request = req.responseText
    Else
        Request = "HTTP Error: " & req.Status & " " & req.statusText
    End If

    Exit Function


ErrorHandler:
    Call Message.Error(err)
    Exit Function
End Function


Function Post(ByVal url As String, ByRef data As String) As String
    On Error GoTo ErrorHandler
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    req.Open "POST", url, False
    req.SetTimeouts 3000
    req.Send data

    ' 检查状态码
    If req.Status = 200 Then
        Post = http.responseText
    Else
        Message.Error "HTTP 请求失败: " & req.Status & " " & req.statusText
    End If

    Exit Function


ErrorHandler:
    Call Message.Error(err)
    Exit Function
End Function
