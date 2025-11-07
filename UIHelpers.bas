Attribute VB_Name = "UIHelpers"
Option Explicit

'공통 컨트롤 생성(점선/3D 제거 포함)

Public Function AddBtn(frm As Object, name$, cap$, x!, y!, w!, h!) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = frm.Controls.Add("Forms.CommandButton.1", name, True)
    With b
        .Caption = cap
        .Left = x: .Top = y: .Width = w: .Height = h
        .Font.name = FONT_NAME: .Font.size = 9
        .TakeFocusOnClick = False            ' 포커스 점선 방지
        '.SpecialEffect = fmSpecialEffectFlat ' 3D 제거 (플랫)
    End With
    Set AddBtn = b
End Function

Public Function AddLbl(frm As Object, name$, cap$, x!, y!, w!, h!, _
                       Optional bold As Boolean = False, _
                       Optional align As Integer = fmTextAlignCenter) As MSForms.Label
    Dim l As MSForms.Label
    Set l = frm.Controls.Add("Forms.Label.1", name, True)
    With l
        .Caption = cap
        .Left = x: .Top = y: .Width = w: .Height = h
        .BackStyle = fmBackStyleTransparent
        .Font.name = FONT_NAME: .Font.size = 12: .Font.bold = bold
        .TextAlign = align
    End With
    Set AddLbl = l
End Function


