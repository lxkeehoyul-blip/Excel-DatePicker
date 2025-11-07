VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===== 상태 =====
Private mYear As Long, mMonth As Long
Private mTarget As Range

' ===== 헤더 컨트롤 =====
Private mLblMonth As MSForms.Label
Private WithEvents wPrevY As MSForms.CommandButton  ' <<
Attribute wPrevY.VB_VarHelpID = -1
Private WithEvents wPrev  As MSForms.CommandButton  ' <
Attribute wPrev.VB_VarHelpID = -1
Private WithEvents wNext  As MSForms.CommandButton  ' >
Attribute wNext.VB_VarHelpID = -1
Private WithEvents wNextY As MSForms.CommandButton  ' >>
Attribute wNextY.VB_VarHelpID = -1
Private WithEvents wToday As MSForms.CommandButton  ' 오늘
Attribute wToday.VB_VarHelpID = -1
Private WithEvents wClear As MSForms.CommandButton  ' 삭제
Attribute wClear.VB_VarHelpID = -1

' ===== 날짜 버튼 묶음 =====
Private mDayBtns  As Collection  ' clsDayButton 모음(이벤트용)
Private mDayCtrls As Collection  ' 실제 CommandButton 컨트롤 모음

' ===== 외부 초기화 API =====
Public Sub InitTarget(ByVal tgt As Range): Set mTarget = tgt: End Sub
Public Sub InitMonth(ByVal y As Long, ByVal m As Long): mYear = y: mMonth = m: End Sub

' ===== 폼 초기화 =====
Private Sub UserForm_Initialize()
    On Error GoTo Fail

    If mYear = 0 Then mYear = Year(Date)
    If mMonth = 0 Then mMonth = Month(Date)
    Me.Caption = "날짜 선택"

    ' 폼 크기(InsideWidth/Height 차이를 고려해 외곽 크기 설정)
    Dim insideW!, insideH!
    insideW = MARGIN * 2 + GridInsideWidth()
    insideH = MARGIN * 2 + GridInsideHeight()
    Me.Width = insideW + (Me.Width - Me.InsideWidth)
    Me.Height = insideH + (Me.Height - Me.InsideHeight)

    Set mDayBtns = New Collection
    Set mDayCtrls = New Collection

    BuildHeader
    BuildWeekHeader
    BuildDayGrid
    Render mYear, mMonth

    AssertHeaderFits
    Exit Sub
Fail:
    MsgBox "Init#" & Err.Number & " - " & Err.Description, vbExclamation, "CalendarForm"
End Sub

' ===== 헤더 구성: 1행(왼쪽 몰기/오른쪽 몰기) + 2행(오른쪽 정렬) =====
Private Sub BuildHeader()
    Dim y!, gw!: y = MARGIN: gw = GridInsideWidth()

    ' 왼쪽: <<  <
    Set wPrevY = AddBtn(Me, "btnPrevY", "<<", MARGIN, y, BTN_W_ICON, BTN_H)
    Set wPrev = AddBtn(Me, "btnPrev", "<", wPrevY.Left + BTN_W_ICON + GAPX, y, BTN_W_ICON, BTN_H)

    ' 오른쪽: >  >>
    Dim rightStart!: rightStart = MARGIN + gw - (BTN_W_ICON * 2 + GAPX)
    Set wNext = AddBtn(Me, "btnNext", ">", rightStart, y, BTN_W_ICON, BTN_H)
    Set wNextY = AddBtn(Me, "btnNextY", ">>", wNext.Left + BTN_W_ICON + GAPX, y, BTN_W_ICON, BTN_H)

    ' 가운데 라벨: 양쪽 그룹 제외 중앙 폭
    Dim grpW!      ' 한쪽 그룹 폭 = 아이콘 2개 + 그룹 내부 간격 1개
    grpW = (2 * BTN_W_ICON) + GAPX
    Dim lblLeft!, lblW!
    lblLeft = wPrev.Left + BTN_W_ICON + GAPX
    lblW = gw - 2 * (grpW + GAPX)                ' = gridW - 2*(grpW + GAPX)
    Set mLblMonth = AddLbl(Me, "lblMonth", "", lblLeft, y, lblW, BTN_H, True, fmTextAlignCenter)

    ' 2행: 오른쪽 정렬 - [오늘][삭제]
    Dim y2!, row2W!, x2!
    y2 = y + BTN_H + GAPY
    row2W = BTN_W_TEXT * 2 + GAPX
    x2 = MARGIN + gw - row2W
    Set wToday = AddBtn(Me, "btnToday", "오늘", x2, y2, BTN_W_TEXT, BTN_H)
    Set wClear = AddBtn(Me, "btnClear", "삭제", wToday.Left + BTN_W_TEXT + GAPX, y2, BTN_W_TEXT, BTN_H)
End Sub

' ===== 요일 헤더(일~토) =====
Private Sub BuildWeekHeader()
    Dim names: names = Array("일", "월", "화", "수", "목", "금", "토")
    Dim i As Long, x!, y!
    x = MARGIN: y = MARGIN + HEADER_H + GAPY

    For i = 0 To GRID_COLS - 1
        Call AddLbl(Me, "lblW" & i, CStr(names(i)), _
                    x + i * (DAY_W + GAPX), y, DAY_W, WEEK_H, True, fmTextAlignCenter)
    Next
End Sub

' ===== 날짜 버튼 42개(최초 1회만 생성) =====
Private Sub BuildDayGrid()
    Dim r As Long, c As Long, i As Long, x!, y!
    x = MARGIN
    y = MARGIN + HEADER_H + GAPY + WEEK_H + GAPY

    Dim cb As MSForms.CommandButton, w As clsDayButton

    For r = 0 To GRID_ROWS - 1
        For c = 0 To GRID_COLS - 1
            i = r * GRID_COLS + c + 1
            Set cb = AddBtn(Me, "btnDay" & i, "", _
                            x + c * (DAY_W + GAPX), _
                            y + r * (DAY_H + GAPY), _
                            DAY_W, DAY_H)
            cb.Font.size = 9
            mDayCtrls.Add cb

            Set w = New clsDayButton
            Set w.Btn = cb: Set w.Host = Me
            mDayBtns.Add w
        Next
    Next
End Sub

' ===== 화면 동기화(유일한 렌더 진입점) =====
Private Sub Render(ByVal y As Long, ByVal m As Long)
    ' 헤더 라벨
    mLblMonth.Caption = Format(dateSerial(y, m, 1), "yyyy-mm")

    ' 전체 리셋
    Dim maxBtn&, idx&, d&, days&, offset&, first As Date
    maxBtn = GRID_COLS * GRID_ROWS
    first = dateSerial(y, m, 1)
    days = Day(dateSerial(y, m + 1, 0))
    offset = Weekday(first, vbSunday) - 1   ' 일요일=0

    For idx = 1 To maxBtn
        With mDayCtrls(idx)
            .Caption = "": .TAG = "": .Enabled = False
            .BackColor = &H8000000F ' 윈도 기본 컨트롤 배경
        End With
    Next

    ' 해당 월 채우기
    For d = 1 To days
        idx = offset + d
        With mDayCtrls(idx)
            .Caption = CStr(d)
            .TAG = CLng(dateSerial(y, m, d))
            .Enabled = True
            ' 주말 연한 배경(선택적)
            If ((idx - 1) Mod 7) = 0 Then .BackColor = RGB(255, 238, 238) ' Sun
            If ((idx - 1) Mod 7) = 6 Then .BackColor = RGB(235, 242, 255) ' Sat
        End With
    Next
End Sub

' ===== 날짜 선택 → 대상 셀에 기록 =====
Public Sub PickDate(ByVal dateSerial As Long)
    If Not mTarget Is Nothing Then
        mTarget.Value = CDate(dateSerial)
        If mTarget.NumberFormatLocal = "General" Or Len(mTarget.NumberFormatLocal) = 0 Then
            mTarget.NumberFormatLocal = "yyyy-mm-dd"
        End If
    End If
    Unload Me
End Sub

' ===== 이동/오늘/삭제(컨트롤러: 상태 변경 → Render 한 줄) =====
Private Sub wPrev_Click()
    mMonth = mMonth - 1
    If mMonth = 0 Then mMonth = 12: mYear = mYear - 1
    Render mYear, mMonth
End Sub

Private Sub wNext_Click()
    mMonth = mMonth + 1
    If mMonth = 13 Then mMonth = 1: mYear = mYear + 1
    Render mYear, mMonth
End Sub

Private Sub wPrevY_Click()
    mYear = mYear - 1
    Render mYear, mMonth
End Sub

Private Sub wNextY_Click()
    mYear = mYear + 1
    Render mYear, mMonth
End Sub

Private Sub wToday_Click()
    If Not mTarget Is Nothing Then
        mTarget.Value = Date
        If mTarget.NumberFormatLocal = "General" Or Len(mTarget.NumberFormatLocal) = 0 Then
            mTarget.NumberFormatLocal = "yyyy-mm-dd"
        End If
    End If
    Unload Me
End Sub

Private Sub wClear_Click()
    If Not mTarget Is Nothing Then SafeClear mTarget
    Unload Me
End Sub


