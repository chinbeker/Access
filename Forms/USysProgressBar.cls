VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

'@Lang VBA

Private ProgressPiece As Integer
Private CurrentLength As Long
Private Const ProgressSize As Long = 5500

Public Sub Start(Optional ByVal Description As String, Optional ByVal piece As Integer = 100)
    On Error Resume Next
    ProgressPiece = piece
    If Not VBA.IsMissing(Description) And Not VBA.IsNull(Description) Then
        Me.Description.Caption = Description
    Else
        Me.Description.Caption = "系统运行中，请等候......"
    End If
    DoEvents
    Me.Repaint
End Sub

Public Sub Run(ByVal val As Double)
    On Error Resume Next
    If val < 0 Then val = 0
    If val > 100 Then val = 100
    CurrentLength = VBA.CLng(ProgressSize / 100 * val + 100)
    Me.Bar.Width = CurrentLength
    If CurrentLength >= ProgressSize Then Me.TimerInterval = 0
    DoEvents
    Me.Repaint
End Sub

Public Sub Complete()
    On Error Resume Next
    Me.TimerInterval = 0
    DoCmd.Close acForm, "USysProgressBar", acSaveNo
End Sub


Private Sub Form_Load()
    ProgressPiece = 100
    If VBA.IsNull(Me.OpenArgs) Then
        Me.TimerInterval = 0
    ElseIf VBA.CBool(Me.OpenArgs) = True Then
        Me.TimerInterval = 250
    Else
        Me.TimerInterval = 0
    End If
End Sub

Private Sub Form_Timer()
    On Error Resume Next
    CurrentLength = VBA.CLng(CurrentLength + ProgressPiece)
    If CurrentLength >= ProgressSize Then
        Me.TimerInterval = 0
        Me.Bar.Width = ProgressSize
    Else
        Me.Bar.Width = CurrentLength
    End If
    DoEvents
    Me.Repaint
End Sub
