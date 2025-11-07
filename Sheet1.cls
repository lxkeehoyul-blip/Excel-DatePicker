VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


' 시트 더블클릭 시 특정 범위에서만 팝업
' ─ 병합셀 허용, 다중셀 선택 차단
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Const TAG As String = "D3:F3"   ' 허용 범위

    If Intersect(Target, Me.Range(TAG)) Is Nothing Then Exit Sub
    If Target.CountLarge > 1 And Not Target.MergeCells Then Exit Sub

    Cancel = True
    ShowCalendar Target   ' ← 표준 모듈의 Public Sub
End Sub

