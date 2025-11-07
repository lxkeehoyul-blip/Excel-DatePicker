VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'³¯Â¥ ¹öÆ° ÀÌº¥Æ® Ä¸½¶È­

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public Host As Object ' CalendarForm

Private Sub Btn_Click()
    If Len(Btn.TAG) > 0 Then
        Host.PickDate CLng(Btn.TAG)
    End If
End Sub

