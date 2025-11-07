Attribute VB_Name = "CalendarEntry"
Option Explicit

'호출/더블클릭 데모(선택)

' 단독 테스트용: 현재 활성셀 기준으로 폼 실행
Public Sub ShowCalendar(Optional ByVal tgt As Range)
    Dim f As New CalendarForm
    If tgt Is Nothing Then Set tgt = ActiveCell
    f.InitTarget tgt
    f.InitMonth Year(Date), Month(Date)
    f.Show
End Sub

' Alt+F8 에 보이게 하려면(선택)
Public Sub CalendarPopup(): ShowCalendar ActiveCell: End Sub


