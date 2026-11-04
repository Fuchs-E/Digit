Attribute VB_Name = "DigitUndo"
'==============================================================================
' DigitUndo - Digit's own undo history.
'
' Running code clears Excel's native undo stack, so before each change Digit
' saves what it is about to overwrite and can put it back. The history holds the
' last N changes (settings: undoDepth, default 10) for this Excel session.
'
' A range snapshot is a list of blocks. Each block is a rectangular range whose
' cells all share the same number format (and font / fill colour when the change
' touches those). Uniform areas, which are the usual case, cost one block each
' however large they are; mixed areas are split in half recursively until each
' part is uniform. That keeps whole-column selections cheap.
'
'   block = Array(address, numberFormat, fontMode, fontColor, fillMode, fillColor, themeColor, tint,
'                 bold, italic, underline)        ' the last three are Empty when not saved
'     fontMode: 0 not saved, 1 automatic, 2 colour, 3 theme colour (themeColor + tint)
'     fillMode: 0 not saved, 1 no fill,   2 colour
'
' New cells use the theme's text colour (Text 1), not "automatic", so theme
' colours are saved as such; restoring them as plain RGB would detach the cell
' from the workbook theme.
'
' Like DigitFormat, this module has no UI.
'==============================================================================
Option Explicit

Private Const MAX_BLOCKS As Long = 50000      ' beyond this, ask before applying without undo

Private mHistory As Collection                ' of Scripting.Dictionary, newest last
Private mDepth As Long
Private mIncStyle As Boolean                  ' current capture also saves bold / italic / underline

'------------------------------------------------------------------------------
' History
'------------------------------------------------------------------------------

Public Sub Undo_SetDepth(ByVal depth As Long)
    If depth < 1 Then depth = 1
    mDepth = depth
    TrimHistory
End Sub

Public Function Undo_Count() As Long
    If mHistory Is Nothing Then Exit Function
    Undo_Count = mHistory.Count
End Function

' Label of the change that Undo would revert next ("" when empty).
Public Function Undo_PeekLabel() As String
    If Undo_Count() = 0 Then Exit Function
    Undo_PeekLabel = mHistory(mHistory.Count)("label")
End Function

Public Sub Undo_Push(ByVal entry As Scripting.Dictionary)
    If mHistory Is Nothing Then Set mHistory = New Collection
    mHistory.Add entry
    TrimHistory
End Sub

Public Sub Undo_Clear()
    Set mHistory = Nothing
End Sub

' Reverts the most recent change. Returns drOK, drNothingToDo or drWarning.
Public Function Undo_Last(ByRef msg As String) As DigitResult
    Dim entry As Scripting.Dictionary
    If Undo_Count() = 0 Then
        msg = "Digit: nothing to undo."
        Undo_Last = drNothingToDo
        Exit Function
    End If
    Set entry = mHistory(mHistory.Count)
    mHistory.Remove mHistory.Count
    If Undo_Restore(entry, msg) Then
        msg = "Digit: undid " & entry("label") & "."
        Undo_Last = drOK
    Else
        Undo_Last = drWarning
    End If
End Function

Private Sub TrimHistory()
    If mDepth < 1 Then mDepth = 10
    If mHistory Is Nothing Then Exit Sub
    Do While mHistory.Count > mDepth
        mHistory.Remove 1
    Loop
End Sub

'------------------------------------------------------------------------------
' Capture
'------------------------------------------------------------------------------

' Snapshot of tgt before a change. Returns Nothing when it would need more than
' MAX_BLOCKS blocks.
Public Function Undo_CaptureRange(ByVal tgt As Range, ByVal includeFont As Boolean, ByVal includeFill As Boolean, _
                                  ByVal label As String, Optional ByVal includeStyle As Boolean = False) As Scripting.Dictionary
    Dim entry As New Scripting.Dictionary, blocks As New Collection, area As Range

    mIncStyle = includeStyle
    For Each area In tgt.Areas
        If Not SnapRange(area, includeFont, includeFill, blocks) Then Exit Function
    Next area

    entry("kind") = "range"
    entry("label") = label
    Set entry("sheet") = tgt.Worksheet
    entry("sheetName") = tgt.Worksheet.Name
    entry("bookName") = tgt.Worksheet.Parent.Name
    Set entry("blocks") = blocks
    Set Undo_CaptureRange = entry
End Function

' Snapshot of pivot data field number formats before a change.
Public Function Undo_CapturePivot(ByVal pt As PivotTable, ByVal fields As Collection, ByVal label As String) As Scripting.Dictionary
    Dim entry As New Scripting.Dictionary, saved As New Collection, df As PivotField

    For Each df In fields
        saved.Add Array(df.Name, CStr(df.NumberFormat))
    Next df

    entry("kind") = "pivot"
    entry("label") = label
    Set entry("sheet") = pt.Parent
    entry("sheetName") = pt.Parent.Name
    entry("bookName") = pt.Parent.Parent.Name
    entry("pivot") = pt.Name
    Set entry("fields") = saved
    Set Undo_CapturePivot = entry
End Function

' Adds blocks for r. Returns False when the block limit is reached.
Private Function SnapRange(ByVal r As Range, ByVal incFont As Boolean, ByVal incFill As Boolean, _
                           ByVal blocks As Collection) As Boolean
    Dim st As Variant, rowsN As Long, colsN As Long, half As Long

    If blocks.Count >= MAX_BLOCKS Then Exit Function

    If ReadUniform(r, incFont, incFill, st) Then
        blocks.Add st
        SnapRange = True
        Exit Function
    End If

    rowsN = r.Rows.Count
    colsN = r.Columns.Count

    If rowsN = 1 And colsN = 1 Then
        ' A single cell can still read as mixed (rich text with several font
        ' colours); keep its number format and leave the font alone.
        blocks.Add Array(r.Address(False, False), r.NumberFormat, 0, 0, FillMode(r, incFill), FillColor(r, incFill), 0, 0#, _
                         Empty, Empty, Empty)
        SnapRange = True
        Exit Function
    End If

    ' Second half starts from its own first cell. r.Offset(...) would shift the
    ' whole range first, which runs off the sheet for entire rows or columns.
    If colsN > 1 Then
        half = colsN \ 2
        If Not SnapRange(r.Resize(rowsN, half), incFont, incFill, blocks) Then Exit Function
        SnapRange = SnapRange(r.Cells(1, half + 1).Resize(rowsN, colsN - half), incFont, incFill, blocks)
    Else
        half = rowsN \ 2
        If Not SnapRange(r.Resize(half, colsN), incFont, incFill, blocks) Then Exit Function
        SnapRange = SnapRange(r.Cells(half + 1, 1).Resize(rowsN - half, colsN), incFont, incFill, blocks)
    End If
End Function

' True when every cell in r shares the properties being saved; st receives the block.
Private Function ReadUniform(ByVal r As Range, ByVal incFont As Boolean, ByVal incFill As Boolean, ByRef st As Variant) As Boolean
    Dim nf As Variant, fMode As Long, fColor As Variant, iMode As Long, iColor As Variant, v As Variant
    Dim theme As Variant, tint As Variant, b As Variant, it As Variant, u As Variant

    nf = r.NumberFormat
    If IsNull(nf) Then Exit Function

    If incFont Then
        v = r.Font.ColorIndex
        If IsNull(v) Then Exit Function
        If v = xlColorIndexAutomatic Then
            fMode = 1: fColor = 0
        Else
            fColor = r.Font.Color
            If IsNull(fColor) Then Exit Function
            fMode = 2
            ' Theme colour? ThemeColor raises an error for a plain RGB font and is
            ' Null when the range mixes theme colours.
            theme = Empty
            On Error Resume Next
            theme = r.Font.ThemeColor
            tint = r.Font.TintAndShade
            On Error GoTo 0
            If IsNull(theme) Or IsNull(tint) Then Exit Function
            If Not IsEmpty(theme) Then
                If theme >= 1 And theme <= 12 Then fMode = 3
            End If
        End If
    Else
        fColor = 0
    End If
    If fMode <> 3 Then theme = 0: tint = 0#

    If incFill Then
        v = r.Interior.Pattern
        If IsNull(v) Then Exit Function
        If v = xlNone Then
            iMode = 1: iColor = 0
        Else
            iColor = r.Interior.Color
            If IsNull(iColor) Then Exit Function
            iMode = 2
        End If
    Else
        iColor = 0
    End If

    b = Empty: it = Empty: u = Empty
    If mIncStyle Then
        b = r.Font.Bold: it = r.Font.Italic: u = r.Font.Underline
        If IsNull(b) Or IsNull(it) Or IsNull(u) Then Exit Function
    End If

    st = Array(r.Address(False, False), CStr(nf), fMode, CLng(fColor), iMode, CLng(iColor), CLng(theme), CDbl(tint), b, it, u)
    ReadUniform = True
End Function

Private Function FillMode(ByVal r As Range, ByVal incFill As Boolean) As Long
    If Not incFill Then Exit Function
    If r.Interior.Pattern = xlNone Then FillMode = 1 Else FillMode = 2
End Function

Private Function FillColor(ByVal r As Range, ByVal incFill As Boolean) As Long
    If Not incFill Then Exit Function
    FillColor = CLng(r.Interior.Color)
End Function

'------------------------------------------------------------------------------
' Restore
'------------------------------------------------------------------------------

' Puts a snapshot back. Returns False with msg when the workbook, sheet or pivot
' is no longer there or cannot be changed.
Public Function Undo_Restore(ByVal entry As Scripting.Dictionary, ByRef msg As String) As Boolean
    Dim ws As Worksheet, b As Variant, r As Range, pt As PivotTable, f As Variant
    Dim prevScreen As Boolean, probe As String, failed As Long, lastErr As String

    On Error Resume Next
    Set ws = entry("sheet")
    probe = ws.Name
    If Err.Number <> 0 Or ws Is Nothing Then
        msg = "Digit cannot undo """ & entry("label") & """: the sheet """ & entry("sheetName") & _
              """ in " & entry("bookName") & " is no longer open."
        Exit Function
    End If
    On Error GoTo Fail

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    If entry("kind") = "pivot" Then
        Set pt = ws.PivotTables(entry("pivot"))
        For Each f In entry("fields")
            pt.DataFields(f(0)).NumberFormat = f(1)
        Next f
    Else
        ' Carry on past a block that fails (for example part of a merged cell)
        ' so one odd cell does not stop the rest being restored.
        On Error Resume Next
        For Each b In entry("blocks")
            Err.Clear
            Set r = ws.Range(b(0))
            r.NumberFormat = b(1)
            Select Case b(2)
                Case 1: r.Font.ColorIndex = xlColorIndexAutomatic
                Case 2: r.Font.Color = b(3)
                Case 3: r.Font.ThemeColor = b(6): r.Font.TintAndShade = b(7)
            End Select
            Select Case b(4)
                Case 1: r.Interior.ColorIndex = xlColorIndexNone
                Case 2: r.Interior.Color = b(5)
            End Select
            If UBound(b) >= 10 Then
                If Not IsEmpty(b(8)) Then r.Font.Bold = b(8)
                If Not IsEmpty(b(9)) Then r.Font.Italic = b(9)
                If Not IsEmpty(b(10)) Then r.Font.Underline = b(10)
            End If
            If Err.Number <> 0 Then failed = failed + 1: lastErr = Err.Description
        Next b
        On Error GoTo Fail
        If failed > 0 Then
            Application.ScreenUpdating = prevScreen
            msg = "Digit restored """ & entry("label") & """ only in part (" & failed & " block(s) failed: " & lastErr & ")."
            If ws.ProtectContents Then msg = msg & vbCrLf & vbCrLf & "The sheet """ & ws.Name & """ is protected."
            Exit Function
        End If
    End If

    Application.ScreenUpdating = prevScreen
    Undo_Restore = True
    Exit Function

Fail:
    msg = "Digit could not undo """ & entry("label") & """." & vbCrLf & vbCrLf & Err.Description
    If ws.ProtectContents Then msg = msg & vbCrLf & vbCrLf & "The sheet """ & ws.Name & """ is protected."
    Application.ScreenUpdating = prevScreen
End Function
