Attribute VB_Name = "SafeOps"
Option Explicit

'병합/보호시트/이벤트 재진입 안전판

Public Sub SafeClear(ByVal r As Range)
    On Error GoTo Done
    Application.EnableEvents = False
    If r.Parent.ProtectContents Then GoTo Done

    If r.MergeCells Then
        r.MergeArea.ClearContents
    Else
        r.ClearContents
    End If
Done:
    Application.EnableEvents = True
End Sub

