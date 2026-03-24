Attribute VB_Name = "NewMMPP"
Option Explicit

' ============================================================
' NewMMPP — starter module for MMPP-related macros
' Import this .bas into your VBA project (Alt+F11 → File → Import).
' Add your public subs here or call existing modules from RUN_NEWMMPP.
' ============================================================

Public Sub RUN_NEWMMPP()
    Dim prevSU As Boolean
    Dim prevEA As Boolean

    prevSU = Application.ScreenUpdating
    prevEA = Application.EnableEvents

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' New sequence entrypoint (local to this module).
    ' Replace the placeholder calls below with your actual steps.
    Call NEW_SEQ_STEP_01
    Call NEW_SEQ_STEP_02

CleanExit:
    Application.ScreenUpdating = prevSU
    Application.EnableEvents = prevEA
    Exit Sub

CleanFail:
    Application.ScreenUpdating = prevSU
    Application.EnableEvents = prevEA
    MsgBox "RUN_NEWMMPP error: " & Err.Description & " (" & Err.Number & ")", vbCritical
    Err.Clear
End Sub

Private Sub NEW_SEQ_STEP_01()
    ' TODO: implement step 1
End Sub

Private Sub NEW_SEQ_STEP_02()
    ' TODO: implement step 2
End Sub
