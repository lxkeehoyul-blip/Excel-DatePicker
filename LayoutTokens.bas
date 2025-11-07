Attribute VB_Name = "LayoutTokens"
Option Explicit

' 레이아웃/수치 토큰 + 기본 좌표 유틸
' ===== Layout Tokens (한 곳에서만 수치 관리) =====
Public Const GRID_COLS As Long = 7
Public Const GRID_ROWS As Long = 6

Public Const DAY_W     As Single = 36
Public Const DAY_H     As Single = 28
Public Const WEEK_H    As Single = 18

Public Const GAPX      As Single = 8
Public Const GAPY      As Single = 8
Public Const MARGIN    As Single = 12

Public Const HEADER_H  As Single = 56

Public Const BTN_H       As Single = 24
Public Const BTN_W_ICON  As Single = 28   ' <<, <, >, >>
Public Const BTN_W_TEXT  As Single = 52   ' 오늘, 삭제 등

Public Const FONT_NAME   As String = "Arial"

' ===== 넓이 계산 유틸(폼 내부 기준) =====
Public Function GridInsideWidth() As Single
    GridInsideWidth = GRID_COLS * DAY_W + (GRID_COLS - 1) * GAPX
End Function

Public Function GridInsideHeight() As Single
    GridInsideHeight = HEADER_H + GAPY + WEEK_H + GAPY + GRID_ROWS * DAY_H + (GRID_ROWS - 1) * GAPY
End Function

' ===== 검산: 헤더가 전체 넓이를 넘지 않는지 =====
Public Sub AssertHeaderFits()
    Dim gw As Single: gw = GridInsideWidth()
    ' 한쪽 그룹(아이콘 2개 + 그룹 내부 간격 1개)
    Dim grpW As Single: grpW = (2 * BTN_W_ICON) + GAPX
    ' 중앙 라벨 폭 = 전체 - 양쪽 그룹 - 양쪽(그룹 옆 간격)
    Dim labelW As Single: labelW = gw - 2 * (grpW + GAPX)
    If labelW < 0 Then Err.Raise 513, , "Header overflow: labelW < 0"
End Sub

