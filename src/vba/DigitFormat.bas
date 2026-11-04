Attribute VB_Name = "DigitFormat"
'==============================================================================
' DigitFormat - the formatting engine: what to format, which preset comes
' next, and applying it.
'
' This module uses the Excel object model but no UI: no MsgBox, no ribbon, no
' UserForm. It reports back through a DigitResult code and a message, and the
' caller (DigitApp) decides how to show it. The C# add-in can port it method
' for method against the same object model.
'==============================================================================
Option Explicit

Private Const MAX_PIVOT_SCAN As Long = 5000   ' cells checked when a selection spans several pivot data fields

Private mCanon As Scripting.Dictionary       ' raw format -> format as Excel stores it

'------------------------------------------------------------------------------
' Main entry point
'------------------------------------------------------------------------------

' Applies the next preset of grp to the current selection.
'   skipUndo = True applies without saving undo state (used after the user
'   agrees to go ahead with a selection that is too large to snapshot).
Public Function Fmt_ApplyGroup(ByVal grp As Scripting.Dictionary, ByRef msg As String, _
                               Optional ByVal skipUndo As Boolean = False) As DigitResult
    Dim presets As Collection, p As Scripting.Dictionary
    Dim sel As Object, ac As Range, tgt As Range, src As Range
    Dim pt As PivotTable, fields As Collection, df As PivotField
    Dim idx As Long, fontC As Long, fillC As Long, label As String
    Dim entry As Scripting.Dictionary, prevScreen As Boolean, changed As Boolean
    Dim ws As Worksheet

    msg = ""
    prevScreen = Application.ScreenUpdating

    If grp Is Nothing Then
        msg = "That preset group no longer exists. Open Digit Settings to check your presets."
        Fmt_ApplyGroup = drWarning
        Exit Function
    End If
    Set presets = grp("presets")
    If presets.Count = 0 Then
        msg = "The preset group """ & grp("name") & """ has no presets."
        Fmt_ApplyGroup = drWarning
        Exit Function
    End If

    If ActiveWorkbook Is Nothing Then
        msg = "Digit: open a workbook first."
        Fmt_ApplyGroup = drNothingToDo
        Exit Function
    End If

    Set sel = Selection
    If sel Is Nothing Then
        msg = "Digit: select some cells first."
        Fmt_ApplyGroup = drNothingToDo
        Exit Function
    End If
    If TypeName(sel) <> "Range" Then
        msg = "Digit: select cells to format (charts and shapes are left alone)."
        Fmt_ApplyGroup = drNothingToDo
        Exit Function
    End If
    Set ac = ActiveCell
    If ac Is Nothing Then Set ac = sel.Cells(1, 1)

    On Error GoTo Fail

    '--- Pivot table data area: format the DataField(s) so it survives refresh.
    If Fmt_PivotTarget(sel, ac, pt, fields) Then
        Set ws = pt.Parent
        idx = NextPresetIndex(presets, CStr(fields(1).NumberFormat), "", "", False)
        Set p = presets(idx)
        label = grp("name") & " " & idx & "/" & presets.Count

        If Not skipUndo Then Set entry = Undo_CapturePivot(pt, fields, label)

        Application.ScreenUpdating = False
        For Each df In fields
            df.NumberFormat = p("numberFormat")
            changed = True
        Next df
        Application.ScreenUpdating = prevScreen

        If Not entry Is Nothing Then Undo_Push entry
        msg = "Digit: " & label & "  " & p("numberFormat") & "  (pivot field: " & FieldNames(fields) & ")"
        Fmt_ApplyGroup = drOK
        Exit Function
    End If

    '--- Normal range or table column.
    Set tgt = Fmt_RangeTarget(sel, ac)
    Set ws = tgt.Worksheet

    ' The active cell decides the cycle step, even when the selection is mixed.
    Set src = ac
    If ac.MergeCells Then Set src = ac.MergeArea.Cells(1, 1)
    idx = NextPresetIndex(presets, CStr(src.NumberFormat), FontHex(src), FillHex(src), True)
    Set p = presets(idx)
    label = grp("name") & " " & idx & "/" & presets.Count
    fontC = Core_HexToColor(p("fontColor"))
    fillC = Core_HexToColor(p("fillColor"))

    If Not skipUndo Then
        Set entry = Undo_CaptureRange(tgt, fontC >= 0, fillC >= 0, label)
        If entry Is Nothing Then
            msg = "The selection is too large or too varied for Digit to save an undo copy."
            Fmt_ApplyGroup = drUndoTooBig
            Exit Function
        End If
    End If

    ' Whole ranges in one go: one call per property, however many cells/areas.
    Application.ScreenUpdating = False
    changed = True
    tgt.NumberFormat = p("numberFormat")
    If fontC >= 0 Then tgt.Font.Color = fontC
    If fillC >= 0 Then tgt.Interior.Color = fillC
    Application.ScreenUpdating = prevScreen

    If Not entry Is Nothing Then Undo_Push entry
    msg = "Digit: " & label & "  " & p("numberFormat") & "  (" & DescribeTarget(sel, tgt) & ")"
    Fmt_ApplyGroup = drOK
    Exit Function

Fail:
    Dim errNum As Long, errDesc As String, dummy As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1            ' leave error-handling mode so Resume Next below applies
    On Error Resume Next
    ' Put back anything that was half applied.
    If changed And Not entry Is Nothing Then Undo_Restore entry, dummy
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0

    If Not ws Is Nothing Then
        If IsSheetProtected(ws) Then
            msg = "The sheet """ & ws.Name & """ is protected, so Digit cannot change its formatting." & vbCrLf & vbCrLf & _
                  "Unprotect it first (Review > Unprotect Sheet), or protect it with ""Format cells"" allowed."
            Fmt_ApplyGroup = drWarning
            Exit Function
        End If
    End If
    msg = "Digit could not apply the format." & vbCrLf & vbCrLf & errDesc & " (error " & errNum & ")"
    Fmt_ApplyGroup = drError
End Function

'------------------------------------------------------------------------------
' Targeting
'------------------------------------------------------------------------------

' What gets formatted for a normal (non-pivot) selection:
'   * one cell inside a table's data body -> that whole table column
'   * anything else -> the selection itself (all areas)
Public Function Fmt_RangeTarget(ByVal sel As Range, ByVal ac As Range) As Range
    Dim lo As ListObject, col As Long

    If sel.CountLarge = 1 Then
        On Error Resume Next
        Set lo = ac.ListObject
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If Not Application.Intersect(ac, lo.DataBodyRange) Is Nothing Then
                    col = ac.Column - lo.Range.Column + 1
                    Set Fmt_RangeTarget = lo.ListColumns(col).DataBodyRange
                    Exit Function
                End If
            End If
        End If
    End If

    Set Fmt_RangeTarget = sel
End Function

' True when the active cell is in a pivot table's values area. Returns the pivot
' and the distinct data fields covered by the selection (the active cell's
' field first).
Public Function Fmt_PivotTarget(ByVal sel As Range, ByVal ac As Range, ByRef pt As PivotTable, _
                                ByRef fields As Collection) As Boolean
    Dim df As PivotField, body As Range, c As Range, n As Long

    Set df = DataFieldOfCell(ac)
    If df Is Nothing Then Exit Function

    Set pt = Nothing
    On Error Resume Next
    Set pt = ac.PivotTable
    On Error GoTo 0
    If pt Is Nothing Then Exit Function

    Set fields = New Collection
    fields.Add df, df.Name

    If sel.CountLarge > 1 Then
        On Error Resume Next
        Set body = Application.Intersect(sel, pt.DataBodyRange)
        On Error GoTo 0
        If Not body Is Nothing Then
            For Each c In body.Cells
                n = n + 1
                If n > MAX_PIVOT_SCAN Then Exit For
                Set df = DataFieldOfCell(c)
                If Not df Is Nothing Then
                    On Error Resume Next
                    fields.Add df, df.Name      ' duplicate key = already listed
                    Err.Clear
                    On Error GoTo 0
                End If
            Next c
        End If
    End If

    Fmt_PivotTarget = True
End Function

' The data field behind a pivot cell, or Nothing when the cell is not in the
' values area (row/column labels, page fields, blanks, or not a pivot at all).
Private Function DataFieldOfCell(ByVal c As Range) As PivotField
    Dim pc As PivotCell, cellType As Long
    On Error Resume Next
    Set pc = c.PivotCell
    If pc Is Nothing Then Exit Function
    cellType = -1
    cellType = pc.PivotCellType
    Select Case cellType
        Case xlPivotCellValue, xlPivotCellSubtotal, xlPivotCellGrandTotal, xlPivotCellCustomSubtotal
            Set DataFieldOfCell = pc.DataField
        Case xlPivotCellDataField
            Set DataFieldOfCell = pc.PivotField
    End Select
    Err.Clear
End Function

'------------------------------------------------------------------------------
' Cycle helpers
'------------------------------------------------------------------------------

' Works out which preset to apply next from the target's current state.
Private Function NextPresetIndex(ByVal presets As Collection, ByVal curFormat As String, ByVal curFont As String, _
                                 ByVal curFill As String, ByVal useColours As Boolean) As Long
    Dim n As Long, i As Long, p As Scripting.Dictionary
    Dim fmts() As String, canon() As String, fonts() As String, fills() As String

    n = presets.Count
    ReDim fmts(1 To n): ReDim canon(1 To n): ReDim fonts(1 To n): ReDim fills(1 To n)
    For i = 1 To n
        Set p = presets(i)
        fmts(i) = p("numberFormat")
        canon(i) = Fmt_Canonical(p("numberFormat"))
        fonts(i) = p("fontColor")
        fills(i) = p("fillColor")
    Next i

    NextPresetIndex = Core_NextIndex(Core_BestMatch(curFormat, curFont, curFill, fmts, canon, fonts, fills, useColours), n)
End Function

' "#RRGGBB" of a cell's font, or "" when automatic or mixed (rich text).
Private Function FontHex(ByVal c As Range) As String
    Dim v As Variant
    On Error Resume Next
    v = c.Font.ColorIndex
    If IsNull(v) Then Exit Function
    If v = xlColorIndexAutomatic Then Exit Function
    v = c.Font.Color
    If IsNull(v) Then Exit Function
    FontHex = Core_ColorToHex(CLng(v))
End Function

' "#RRGGBB" of a cell's fill, or "" when it has no fill.
Private Function FillHex(ByVal c As Range) As String
    Dim v As Variant
    On Error Resume Next
    v = c.Interior.Pattern
    If IsNull(v) Then Exit Function
    If v = xlNone Then Exit Function
    v = c.Interior.Color
    If IsNull(v) Then Exit Function
    FillHex = Core_ColorToHex(CLng(v))
End Function

'------------------------------------------------------------------------------
' Number format utilities (also used by the settings window)
'
' They use a scratch cell on the add-in's own hidden sheet, so the user's
' workbooks are never touched.
'------------------------------------------------------------------------------

' The format string as Excel stores it once applied (Excel may rewrite it, for
' example "$#,##0" can come back as "\$#,##0"). Cached per format.
Public Function Fmt_Canonical(ByVal fmt As String) As String
    Dim cell As Range, wasSaved As Boolean
    If mCanon Is Nothing Then Set mCanon = New Scripting.Dictionary
    If mCanon.Exists(fmt) Then
        Fmt_Canonical = mCanon(fmt)
        Exit Function
    End If

    Fmt_Canonical = fmt
    wasSaved = ThisWorkbook.Saved
    On Error GoTo Done
    Set cell = ScratchCell()
    cell.NumberFormat = fmt
    Fmt_Canonical = cell.NumberFormat
    cell.NumberFormat = "General"
Done:
    ThisWorkbook.Saved = wasSaved
    mCanon(fmt) = Fmt_Canonical
End Function

Public Sub Fmt_ClearCache()
    Set mCanon = Nothing
End Sub

' True when Excel accepts the number format string.
Public Function Fmt_IsValidNumberFormat(ByVal fmt As String) As Boolean
    Dim cell As Range, wasSaved As Boolean
    If Len(Trim$(fmt)) = 0 Then Exit Function
    wasSaved = ThisWorkbook.Saved
    On Error GoTo Done
    Set cell = ScratchCell()
    cell.NumberFormat = fmt
    Fmt_IsValidNumberFormat = True
    cell.NumberFormat = "General"
Done:
    ThisWorkbook.Saved = wasSaved
End Function

' How value looks in the given format, as Excel would display it.
Public Function Fmt_Preview(ByVal fmt As String, ByVal value As Double) As String
    Dim cell As Range, wasSaved As Boolean, s As String
    wasSaved = ThisWorkbook.Saved
    On Error GoTo Bad
    Set cell = ScratchCell()
    cell.NumberFormat = fmt
    cell.Value = value
    s = cell.Text
    cell.ClearContents
    cell.NumberFormat = "General"
    If Len(s) > 0 And Len(Replace(s, "#", "")) = 0 Then s = "(cannot display this value)"
    Fmt_Preview = s
    ThisWorkbook.Saved = wasSaved
    Exit Function
Bad:
    Fmt_Preview = "(invalid format)"
    ThisWorkbook.Saved = wasSaved
End Function

Private Function ScratchCell() As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    If ws.Columns(1).ColumnWidth < 200 Then ws.Columns(1).ColumnWidth = 200
    Set ScratchCell = ws.Range("A1")
End Function

'------------------------------------------------------------------------------
' Misc
'------------------------------------------------------------------------------

Private Function IsSheetProtected(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    IsSheetProtected = ws.ProtectContents
End Function

Private Function FieldNames(ByVal fields As Collection) As String
    Dim df As PivotField, s As String
    For Each df In fields
        If Len(s) > 0 Then s = s & ", "
        s = s & df.Name
    Next df
    FieldNames = s
End Function

Private Function DescribeTarget(ByVal sel As Range, ByVal tgt As Range) As String
    If tgt.Address(External:=True) <> sel.Address(External:=True) Then
        DescribeTarget = "table column " & tgt.Address(False, False)
    ElseIf tgt.Areas.Count > 1 Then
        DescribeTarget = tgt.Areas.Count & " areas"
    Else
        DescribeTarget = tgt.Address(False, False)
    End If
End Function
