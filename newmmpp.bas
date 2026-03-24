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
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsNewScProd As Worksheet
    Dim wsNewSf As Worksheet
    Dim lastRow As Long
    Dim alertState As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Err.Raise vbObjectError + 2101, "NEW_SEQ_STEP_01", "No active workbook."
    End If

    On Error Resume Next
    Set wsSource = wb.Worksheets("SC_PROD")
    On Error GoTo 0
    If wsSource Is Nothing Then
        Err.Raise vbObjectError + 2102, "NEW_SEQ_STEP_01", "Source sheet not found: SC_PROD"
    End If

    alertState = Application.DisplayAlerts

    On Error Resume Next
    Set wsNewScProd = wb.Worksheets("NEW_SC_PROD")
    On Error GoTo 0
    If Not wsNewScProd Is Nothing Then
        Application.DisplayAlerts = False
        wsNewScProd.Delete
        Application.DisplayAlerts = alertState
        Set wsNewScProd = Nothing
    End If

    wsSource.Copy After:=wsSource
    Set wsNewScProd = wb.ActiveSheet
    wsNewScProd.Name = "NEW_SC_PROD"

    wsNewScProd.Columns("H:H").Insert Shift:=xlToRight
    wsNewScProd.Range("H1").Value = "ConcatCode"

    lastRow = wsNewScProd.Cells(wsNewScProd.Rows.Count, "G").End(xlUp).Row
    If lastRow >= 2 Then
        wsNewScProd.Range("H2:H" & lastRow).FormulaR1C1 = "=CONCATENATE(RC[-2],RC[-1])"
    End If

    On Error Resume Next
    Set wsNewSf = wb.Worksheets("NEW_SF")
    On Error GoTo 0
    If Not wsNewSf Is Nothing Then
        Application.DisplayAlerts = False
        wsNewSf.Delete
        Application.DisplayAlerts = alertState
        Set wsNewSf = Nothing
    End If

    Set wsNewSf = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsNewSf.Name = "NEW_SF"
    wsNewSf.Range("A1").Value = "ConcatCode"

    lastRow = wsNewScProd.Cells(wsNewScProd.Rows.Count, "H").End(xlUp).Row
    wsNewSf.Range("A1:A" & lastRow).Value = wsNewScProd.Range("H1:H" & lastRow).Value

    Application.DisplayAlerts = alertState
End Sub

Private Sub NEW_SEQ_STEP_02()
    ' Reserved for next sequence steps.
End Sub
