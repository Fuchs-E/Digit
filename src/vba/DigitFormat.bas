Attribute VB_Name = "DigitFormat"
'==============================================================================
' DigitFormat - the formatting engine: what to format, which preset comes
' next, and applying it. Also Colour by content (Fmt_ColourByContent).
'
' This module uses the Excel object model but no UI: no MsgBox, no ribbon, no
' UserForm. It reports back through a DigitResult code and a message, and the
' caller (DigitApp) decides how to show it. The C# add-in can port it method
' for method against the same object model.
'==============================================================================
Option Explicit

Private Const MAX_PIVOT_SCAN As Long = 5000   ' cells checked when a selection spans several pivot data fields
Private Const MAX_SPILL_SCAN As Long = 5000   ' formula cells checked for dynamic array spills (Colour by content)
Private Const CC_CHUNK_AREAS As Long = 500    ' areas per range when collecting cells to colour

Private mCanon As Scripting.Dictionary       ' raw format -> format as Excel stores it

' Colour by content, while one run is in progress: the cells of each kind
' (index = DigitContentKind), as finished ranges, the range being built, and
' the cell count.
Private mCcParts(1 To DIGIT_CONTENT_KINDS) As Collection
Private mCcPending(1 To DIGIT_CONTENT_KINDS) As Range
Private mCcCount(1 To DIGIT_CONTENT_KINDS) As Double

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
    Dim idx As Long, fontC As Long, fillC As Long, label As String, style As String
    Dim entry As Scripting.Dictionary, prevScreen As Boolean, changed As Boolean
    Dim ws As Worksheet, stage As String

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

    stage = "checking for a pivot table"
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
    stage = "working out which cells to format"
    Set tgt = Fmt_RangeTarget(sel, ac)
    Set ws = tgt.Worksheet

    ' The active cell decides the cycle step, even when the selection is mixed.
    stage = "reading the active cell's format"
    Set src = ac
    If ac.MergeCells Then Set src = ac.MergeArea.Cells(1, 1)
    idx = NextPresetIndex(presets, CStr(src.NumberFormat), FontHex(src), FillHex(src), True, CellStyle(src))
    Set p = presets(idx)
    label = grp("name") & " " & idx & "/" & presets.Count
    fontC = Core_HexToColor(p("fontColor"))
    fillC = Core_HexToColor(p("fillColor"))
    style = Store_PresetStyle(p)

    If Not skipUndo Then
        stage = "saving an undo copy"
        Set entry = Undo_CaptureRange(tgt, fontC >= 0, fillC >= 0, label, style <> "---")
        If entry Is Nothing Then
            msg = "The selection is too large or too varied for Digit to save an undo copy."
            Fmt_ApplyGroup = drUndoTooBig
            Exit Function
        End If
    End If

    ' Whole ranges in one go: one call per property, however many cells/areas.
    Application.ScreenUpdating = False
    changed = True
    stage = "setting the number format"
    tgt.NumberFormat = p("numberFormat")
    stage = "setting the font colour"
    If fontC >= 0 Then tgt.Font.Color = fontC
    stage = "setting the fill colour"
    If fillC >= 0 Then tgt.Interior.Color = fillC
    stage = "setting bold, italic and underline"
    If Mid$(style, 1, 1) <> "-" Then tgt.Font.Bold = (Mid$(style, 1, 1) = "1")
    If Mid$(style, 2, 1) <> "-" Then tgt.Font.Italic = (Mid$(style, 2, 1) = "1")
    If Mid$(style, 3, 1) = "1" Then tgt.Font.Underline = xlUnderlineStyleSingle
    If Mid$(style, 3, 1) = "0" Then tgt.Font.Underline = xlUnderlineStyleNone
    Application.ScreenUpdating = prevScreen

    If Not entry Is Nothing Then Undo_Push entry
    stage = "writing the status message"
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
    msg = "Digit could not apply the format." & vbCrLf & vbCrLf & errDesc & " (error " & errNum & _
          IIf(Len(stage) > 0, ", while " & stage, "") & ")"
    Fmt_ApplyGroup = drError
End Function

'------------------------------------------------------------------------------
' Colour by content
'------------------------------------------------------------------------------

' Sets the font colour of every cell in the selection by what it contains:
' typed numbers (inputs), formulas, links to other sheets and links to other
' workbooks, in the colours from Settings. Nothing else about a cell changes.
' Text and blank cells, pivot tables, and kinds whose colour is blank are left
' alone. Same result codes and skipUndo meaning as Fmt_ApplyGroup.
Public Function Fmt_ColourByContent(ByRef msg As String, Optional ByVal skipUndo As Boolean = False) As DigitResult
    Dim sel As Object, ws As Worksheet, used As Range, consts As Range, forms As Range
    Dim cfg As Scripting.Dictionary, colour(1 To DIGIT_CONTENT_KINDS) As Long, anyColour As Boolean
    Dim names As Scripting.Dictionary, sheets As Scripting.Dictionary, cache As New Scripting.Dictionary
    Dim area As Range, cell As Range, piece As Variant, spill As Range, counted As Range, allParts As New Collection
    Dim entry As Scripting.Dictionary, prevScreen As Boolean, changed As Boolean, stage As String
    Dim k As Long

    msg = ""
    prevScreen = Application.ScreenUpdating

    If ActiveWorkbook Is Nothing Then
        msg = "Digit: open a workbook first."
        Fmt_ColourByContent = drNothingToDo
        Exit Function
    End If
    Set sel = Selection
    If sel Is Nothing Then
        msg = "Digit: select some cells first."
        Fmt_ColourByContent = drNothingToDo
        Exit Function
    End If
    If TypeName(sel) <> "Range" Then
        msg = "Digit: select cells to colour (charts and shapes are left alone)."
        Fmt_ColourByContent = drNothingToDo
        Exit Function
    End If

    Set cfg = Store_Config()
    For k = ckInput To ckExternal
        colour(k) = Core_HexToColor(Store_ContentColor(cfg, k))
        If colour(k) >= 0 Then anyColour = True
    Next k
    If Not anyColour Then
        msg = "Every Colour by content colour is blank, so there is nothing to apply." & vbCrLf & vbCrLf & _
              "Choose the colours in Digit Settings, on the Colour by content tab."
        Fmt_ColourByContent = drWarning
        Exit Function
    End If

    On Error GoTo Fail
    Set ws = sel.Worksheet
    Cc_Reset

    ' Only cells with content matter, so a whole column or sheet stays cheap.
    stage = "finding the cells with content"
    Set used = Application.Intersect(sel, ws.UsedRange)
    If Not used Is Nothing Then
        Set consts = OutsidePivots(ContentCells(used, False), ws)
        Set forms = OutsidePivots(ContentCells(used, True), ws)
    End If

    stage = "reading the formulas"
    If Not consts Is Nothing Then
        If colour(ckInput) >= 0 Then Cc_Add ckInput, consts
    End If
    If Not forms Is Nothing Then
        Cc_Lookups ws.Parent, names, sheets
        For Each area In forms.Areas
            Cc_AddFormulaArea area, ws.Name, names, sheets, cache, colour
        Next area

        ' A dynamic array formula colours the cells it spills into as well.
        If forms.CountLarge <= MAX_SPILL_SCAN Then
            For Each cell In forms.Cells
                Set spill = SpillOf(cell, sel)
                If Not spill Is Nothing Then
                    k = Cc_Kind(CStr(cell.FormulaR1C1), ws.Name, names, sheets, cache)
                    If colour(k) >= 0 Then
                        Cc_Add k, spill
                        ' Do not count twice the cells already listed as formulas.
                        Set counted = Application.Intersect(spill, forms)
                        If Not counted Is Nothing Then mCcCount(k) = mCcCount(k) - counted.CountLarge
                    End If
                End If
            Next cell
        End If
    End If

    For k = ckInput To ckExternal
        Cc_Flush k
        For Each piece In mCcParts(k)
            allParts.Add piece
        Next piece
    Next k

    If allParts.Count = 0 Then
        Cc_Reset
        If IsSheetProtected(ws) Then
            ' Excel will not list the formula cells of a protected sheet.
            msg = "The sheet """ & ws.Name & """ is protected, so Digit cannot colour its cells by content." & vbCrLf & vbCrLf & _
                  "Unprotect it first (Review > Unprotect Sheet)."
            Fmt_ColourByContent = drWarning
        Else
            msg = "Digit: nothing to colour in the selection (no numbers or formulas)."
            Fmt_ColourByContent = drNothingToDo
        End If
        Exit Function
    End If

    If Not skipUndo Then
        stage = "saving an undo copy"
        Set entry = Undo_CaptureFont(allParts, "Colour by content")
        If entry Is Nothing Then
            Cc_Reset
            msg = "The selection is too large or too varied for Digit to save an undo copy."
            Fmt_ColourByContent = drUndoTooBig
            Exit Function
        End If
    End If

    stage = "setting the font colours"
    Application.ScreenUpdating = False
    changed = True
    For k = ckInput To ckExternal
        For Each piece In mCcParts(k)
            piece.Font.Color = colour(k)
        Next piece
    Next k
    Application.ScreenUpdating = prevScreen

    If Not entry Is Nothing Then Undo_Push entry
    msg = "Digit: Colour by content  " & Cc_Summary()
    Cc_Reset
    Fmt_ColourByContent = drOK
    Exit Function

Fail:
    Dim errNum As Long, errDesc As String, dummy As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If changed And Not entry Is Nothing Then Undo_Restore entry, dummy
    Application.ScreenUpdating = prevScreen
    Cc_Reset
    On Error GoTo 0

    If Not ws Is Nothing Then
        If IsSheetProtected(ws) Then
            msg = "The sheet """ & ws.Name & """ is protected, so Digit cannot change its formatting." & vbCrLf & vbCrLf & _
                  "Unprotect it first (Review > Unprotect Sheet), or protect it with ""Format cells"" allowed."
            Fmt_ColourByContent = drWarning
            Exit Function
        End If
    End If
    msg = "Digit could not colour the cells." & vbCrLf & vbCrLf & errDesc & " (error " & errNum & _
          IIf(Len(stage) > 0, ", while " & stage, "") & ")"
    Fmt_ColourByContent = drError
End Function

' The cells of r that hold a formula (formulas = True) or a typed number, date,
' TRUE/FALSE or error value (formulas = False). Nothing when there are none.
Private Function ContentCells(ByVal r As Range, ByVal formulas As Boolean) As Range
    Dim found As Range, v As Variant
    On Error Resume Next
    If r.CountLarge = 1 Then
        ' SpecialCells on a single cell searches the whole sheet instead.
        If CBool(r.HasFormula) = formulas Then
            If formulas Then
                Set found = r
            Else
                v = r.Value
                If Not IsEmpty(v) And VarType(v) <> vbString Then Set found = r
            End If
        End If
    ElseIf formulas Then
        Set found = r.SpecialCells(xlCellTypeFormulas)
    Else
        Set found = r.SpecialCells(xlCellTypeConstants, xlNumbers + xlLogical + xlErrors)
    End If
    ' Never beyond r, whatever SpecialCells searched (a lone merged cell counts
    ' as one cell to it).
    If Not found Is Nothing Then Set found = Application.Intersect(found, r)
    Set ContentCells = found
End Function

' r without any cell that lies inside a pivot table.
Private Function OutsidePivots(ByVal r As Range, ByVal ws As Worksheet) As Range
    Dim pt As PivotTable, pivots As Range, area As Range, c As Range, keep As Range, n As Long

    If r Is Nothing Then Exit Function
    Set OutsidePivots = r
    On Error Resume Next
    n = ws.PivotTables.Count
    If n = 0 Then Exit Function
    For Each pt In ws.PivotTables
        If pivots Is Nothing Then
            Set pivots = pt.TableRange2
        Else
            Set pivots = Application.Union(pivots, pt.TableRange2)
        End If
    Next pt
    If pivots Is Nothing Then Exit Function
    If Application.Intersect(r, pivots) Is Nothing Then Exit Function

    For Each area In r.Areas
        If Application.Intersect(area, pivots) Is Nothing Then
            Set keep = JoinRanges(keep, area)
        Else
            For Each c In area.Cells
                If Application.Intersect(c, pivots) Is Nothing Then Set keep = JoinRanges(keep, c)
            Next c
        End If
    Next area
    Set OutsidePivots = keep
End Function

Private Function JoinRanges(ByVal a As Range, ByVal b As Range) As Range
    If a Is Nothing Then
        Set JoinRanges = b
    Else
        Set JoinRanges = Application.Union(a, b)
    End If
End Function

' The part of a dynamic array formula's spill range inside sel, or Nothing
' when cell does not spill (or this Excel has no dynamic arrays).
Private Function SpillOf(ByVal cell As Object, ByVal sel As Range) As Range
    Dim sp As Range
    On Error Resume Next
    Set sp = cell.SpillingToRange
    If sp Is Nothing Then Exit Function
    If sp.CountLarge < 2 Then Exit Function
    Set SpillOf = Application.Intersect(sp, sel)
End Function

' What the classifier needs to know about the workbook: where each defined
' name and table points, and the sheet names.
Private Sub Cc_Lookups(ByVal wb As Workbook, ByRef names As Scripting.Dictionary, ByRef sheets As Scripting.Dictionary)
    Dim sh As Object, nm As Object, lo As ListObject, key As String, target As String, p As Long, q As String

    Set names = New Scripting.Dictionary
    Set sheets = New Scripting.Dictionary
    On Error Resume Next

    For Each sh In wb.Sheets
        sheets(LCase$(sh.Name)) = True
    Next sh
    ' The workbook's own name: Excel writes Book.xlsx!Name for a global name
    ' that a sheet-level name hides. That is not a link to another file.
    sheets(LCase$(wb.Name)) = "workbook"

    For Each nm In wb.Names
        key = ""
        target = ""
        key = nm.Name
        target = nm.RefersTo
        If Len(key) > 0 And Left$(target, 1) = "=" Then
            p = InStrRev(key, "!")
            If p > 0 Then
                ' A sheet-scoped name reads Sheet1!Name or 'My Sheet'!Name.
                q = Left$(key, p - 1)
                If Left$(q, 1) = "'" And Right$(q, 1) = "'" And Len(q) >= 2 Then
                    q = Replace(Mid$(q, 2, Len(q) - 2), "''", "'")
                End If
                key = LCase$(q) & "!" & LCase$(Mid$(key, p + 1))
            Else
                key = LCase$(key)
            End If
            names(key) = target
        End If
    Next nm

    For Each sh In wb.Worksheets
        For Each lo In sh.ListObjects
            names(LCase$(lo.Name)) = "='" & Replace(sh.Name, "'", "''") & "'!A1"
        Next lo
    Next sh
End Sub

' Kind of one formula, remembered per distinct text. In R1C1 notation a
' filled-down column is one text, so a large block costs one classification.
Private Function Cc_Kind(ByVal formulaText As String, ByVal sheetName As String, ByVal names As Scripting.Dictionary, _
                         ByVal sheets As Scripting.Dictionary, ByVal cache As Scripting.Dictionary) As Long
    Dim kind As Long
    If cache.Exists(formulaText) Then
        Cc_Kind = cache(formulaText)
        Exit Function
    End If
    kind = Core_ClassifyFormula(formulaText, sheetName, names, sheets)
    If kind < ckInput Then kind = ckFormula
    cache(formulaText) = kind
    Cc_Kind = kind
End Function

' Sorts one rectangular block of formula cells into the kinds. Cells of the
' same kind that sit next to each other are added as one piece.
Private Sub Cc_AddFormulaArea(ByVal area As Range, ByVal sheetName As String, ByVal names As Scripting.Dictionary, _
                              ByVal sheets As Scripting.Dictionary, ByVal cache As Scripting.Dictionary, _
                              ByRef colour() As Long)
    Dim v As Variant, rowsN As Long, colsN As Long, r As Long, c As Long
    Dim kinds() As Byte, k As Long, first As Long, uniform As Boolean, startAt As Long

    If area.CountLarge = 1 Then
        k = Cc_Kind(CStr(area.FormulaR1C1), sheetName, names, sheets, cache)
        If colour(k) >= 0 Then Cc_Add k, area
        Exit Sub
    End If

    v = area.FormulaR1C1
    rowsN = UBound(v, 1)
    colsN = UBound(v, 2)
    ReDim kinds(1 To rowsN, 1 To colsN)
    uniform = True
    For r = 1 To rowsN
        For c = 1 To colsN
            k = Cc_Kind(CStr(v(r, c)), sheetName, names, sheets, cache)
            kinds(r, c) = k
            If first = 0 Then first = k
            If k <> first Then uniform = False
        Next c
    Next r

    If uniform Then
        If colour(first) >= 0 Then Cc_Add first, area
        Exit Sub
    End If

    If rowsN >= colsN Then
        ' Runs down each column.
        For c = 1 To colsN
            startAt = 1
            For r = 2 To rowsN + 1
                If r > rowsN Then
                    k = 0
                Else
                    k = kinds(r, c)
                End If
                If k <> kinds(startAt, c) Then
                    If colour(kinds(startAt, c)) >= 0 Then Cc_Add kinds(startAt, c), area.Cells(startAt, c).Resize(r - startAt, 1)
                    startAt = r
                End If
            Next r
        Next c
    Else
        ' Runs along each row.
        For r = 1 To rowsN
            startAt = 1
            For c = 2 To colsN + 1
                If c > colsN Then
                    k = 0
                Else
                    k = kinds(r, c)
                End If
                If k <> kinds(r, startAt) Then
                    If colour(kinds(r, startAt)) >= 0 Then Cc_Add kinds(r, startAt), area.Cells(r, startAt).Resize(1, c - startAt)
                    startAt = c
                End If
            Next c
        Next r
    End If
End Sub

' Adds a piece to a kind. Pieces are joined into ranges of at most
' CC_CHUNK_AREAS areas, because a union of very many areas gets slow.
Private Sub Cc_Add(ByVal kind As Long, ByVal piece As Range)
    mCcCount(kind) = mCcCount(kind) + piece.CountLarge
    If mCcPending(kind) Is Nothing Then
        Set mCcPending(kind) = piece
    Else
        Set mCcPending(kind) = Application.Union(mCcPending(kind), piece)
    End If
    If mCcPending(kind).Areas.Count >= CC_CHUNK_AREAS Then Cc_Flush kind
End Sub

Private Sub Cc_Flush(ByVal kind As Long)
    If mCcPending(kind) Is Nothing Then Exit Sub
    mCcParts(kind).Add mCcPending(kind)
    Set mCcPending(kind) = Nothing
End Sub

Private Sub Cc_Reset()
    Dim k As Long
    For k = 1 To DIGIT_CONTENT_KINDS
        Set mCcParts(k) = New Collection
        Set mCcPending(k) = Nothing
        mCcCount(k) = 0
    Next k
End Sub

' "12 input, 30 formula, 4 sheet link, 1 external" (kinds with no cells left out).
Private Function Cc_Summary() As String
    Dim k As Long, s As String, label As Variant
    label = Array("", "input", "formula", "sheet link", "external")
    For k = ckInput To ckExternal
        If mCcCount(k) > 0 Then
            If Len(s) > 0 Then s = s & ", "
            s = s & Format$(mCcCount(k), "#,##0") & " " & label(k)
        End If
    Next k
    Cc_Summary = s
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
                                 ByVal curFill As String, ByVal useColours As Boolean, _
                                 Optional ByVal curStyle As String = "???") As Long
    Dim n As Long, i As Long, p As Scripting.Dictionary
    Dim fmts() As String, canon() As String, fonts() As String, fills() As String, styles() As String

    n = presets.Count
    ReDim fmts(1 To n): ReDim canon(1 To n): ReDim fonts(1 To n): ReDim fills(1 To n): ReDim styles(1 To n)
    For i = 1 To n
        Set p = presets(i)
        fmts(i) = p("numberFormat")
        canon(i) = Fmt_Canonical(p("numberFormat"))
        fonts(i) = p("fontColor")
        fills(i) = p("fillColor")
        styles(i) = Store_PresetStyle(p)
    Next i

    NextPresetIndex = Core_NextIndex(Core_BestMatch(curFormat, curFont, curFill, fmts, canon, fonts, fills, useColours, _
                                                    curStyle, styles), n)
End Function

' A cell's bold, italic, underline as "1" / "0" ("?" when mixed, e.g. rich text).
Private Function CellStyle(ByVal c As Range) As String
    Dim v As Variant, s As String
    On Error Resume Next
    v = c.Font.Bold
    If IsNull(v) Then s = "?" Else s = IIf(v, "1", "0")
    v = c.Font.Italic
    If IsNull(v) Then s = s & "?" Else s = s & IIf(v, "1", "0")
    v = c.Font.Underline
    If IsNull(v) Then
        s = s & "?"
    ElseIf v = xlUnderlineStyleNone Then
        s = s & "0"
    Else
        s = s & "1"
    End If
    CellStyle = s
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
