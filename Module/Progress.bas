Attribute VB_Name = "Progress"
'@Lang VBA

Option Compare Database
Option Explicit

Private ProgressBar As Form

Public Sub Start(Optional ByVal Description As String, Optional ByVal auto As Boolean = False, Optional ByVal piece As Integer = 100)
    On Error Resume Next
    If ProgressBar Is Nothing Then
        DoCmd.OpenForm "USysProgressBar", , , , , , auto
        Set ProgressBar = Application.Forms("USysProgressBar")
        If Not ProgressBar Is Nothing Then Call ProgressBar.Start(Description, piece)
    End If
End Sub

Public Sub Run(ByVal val As Double)
    On Error Resume Next
    If Not ProgressBar Is Nothing Then
        Call ProgressBar.Run(val)
        If val >= 100 Then Set ProgressBar = Nothing
    End If
End Sub

Public Sub Complete()
    On Error Resume Next
    If Not ProgressBar Is Nothing Then
        Call ProgressBar.Complete
        Set ProgressBar = Nothing
    Else
        Core.CloseForm "USysProgressBar"
    End If
End Sub
