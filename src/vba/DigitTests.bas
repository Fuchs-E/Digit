Attribute VB_Name = "DigitTests"
'==============================================================================
' DigitTests - self-tests that run inside Excel.
'
'   build.ps1 -RunTests   runs them headless after building and prints the result
'   Digit_RunTests        run from the Immediate window / Alt+F8 for a message box
'
' Unit tests cover DigitCore and DigitStore. Integration tests drive the engine
' and undo against a scratch workbook that is closed without saving. They never
' touch your presets file (a temporary one is used) or your open workbooks.
'==============================================================================
Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As String

Public Sub Digit_RunTests()
    MsgBox RunAllTests(), vbInformation, "Digit tests"
End Sub

' Returns a summary; the first line is "PASS n" or "FAIL n of m".
Public Function RunAllTests(Optional ByVal tempFolder As String = "") As String
    Dim prevAlerts As Boolean
    mPass = 0: mFail = 0: mLog = ""

    If Len(tempFolder) = 0 Then tempFolder = Environ$("TEMP")
    Store_SetPathOverride tempFolder & "\digit-test-" & Format$(Now, "yyyymmddhhnnss") & "\presets.json"

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Crash

    Test_NextIndex
    Test_BestMatch
    Test_Colors
    Test_Shortcuts
    Test_XmlEscape
    Test_Versions
    Test_Collections
    Test_Normalize
    Test_StoreRoundTrip
    Test_FormatUtilities
    Test_RangeCycleAndUndo
    Test_MultiAreaAndColours
    Test_TableColumn
    Test_ProtectedSheet
    Test_PivotDataField
    Test_Icons
    Test_SettingsForm
    GoTo Finish

Crash:
    Fail "unexpected error: " & Err.Description & " (" & Err.Number & ")"

Finish:
    On Error Resume Next
    Undo_Clear
    Store_SetPathOverride ""
    Fmt_ClearCache
    Application.DisplayAlerts = prevAlerts
    If mFail = 0 Then
        RunAllTests = "PASS " & mPass & vbCrLf & mLog
    Else
        RunAllTests = "FAIL " & mFail & " of " & (mPass + mFail) & vbCrLf & mLog
    End If
End Function

'------------------------------------------------------------------------------
' Assertions
'------------------------------------------------------------------------------

Private Sub Check(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        mPass = mPass + 1
    Else
        Fail testName
    End If
End Sub

Private Sub CheckEq(ByVal actual As Variant, ByVal expected As Variant, ByVal testName As String)
    If actual = expected Then
        mPass = mPass + 1
    Else
        Fail testName & ": expected [" & expected & "] got [" & actual & "]"
    End If
End Sub

Private Sub Fail(ByVal text As String)
    mFail = mFail + 1
    mLog = mLog & "  FAIL " & text & vbCrLf
End Sub

'------------------------------------------------------------------------------
' Unit tests
'------------------------------------------------------------------------------

Private Sub Test_NextIndex()
    CheckEq Core_NextIndex(0, 3), 1, "no match -> 1"
    CheckEq Core_NextIndex(1, 3), 2, "1 -> 2"
    CheckEq Core_NextIndex(2, 3), 3, "2 -> 3"
    CheckEq Core_NextIndex(3, 3), 1, "last loops to 1"
    CheckEq Core_NextIndex(1, 2), 2, "two presets 1 -> 2"
    CheckEq Core_NextIndex(2, 2), 1, "two presets loop"
    CheckEq Core_NextIndex(1, 1), 1, "single preset stays"
    CheckEq Core_NextIndex(0, 1), 1, "single preset no match"
End Sub

Private Sub Test_BestMatch()
    Dim f(1 To 3) As String, c(1 To 3) As String, fo(1 To 3) As String, fi(1 To 3) As String
    f(1) = "0%": f(2) = "0.0%": f(3) = "0.00%"
    c(1) = "0%": c(2) = "0.0%": c(3) = "0.00%"
    CheckEq Core_BestMatch("0.0%", "", "", f, c, fo, fi, True), 2, "match by format"
    CheckEq Core_BestMatch("General", "", "", f, c, fo, fi, True), 0, "no match"
    CheckEq Core_BestMatch("0.00%", "#FF0000", "", f, c, fo, fi, True), 3, "blank colours are wildcards"

    ' Same number format, told apart by colour: the most specific match wins.
    f(1) = "#,##0": f(2) = "#,##0": f(3) = "#,##0"
    c(1) = "#,##0": c(2) = "#,##0": c(3) = "#,##0"
    fo(1) = "": fo(2) = "#0000FF": fo(3) = "#0000FF"
    fi(1) = "": fi(2) = "": fi(3) = "#FFF2CC"
    CheckEq Core_BestMatch("#,##0", "", "", f, c, fo, fi, True), 1, "plain"
    CheckEq Core_BestMatch("#,##0", "#0000FF", "", f, c, fo, fi, True), 2, "blue"
    CheckEq Core_BestMatch("#,##0", "#0000ff", "#fff2cc", f, c, fo, fi, True), 3, "blue + fill (case-insensitive)"
    CheckEq Core_BestMatch("#,##0", "#0000FF", "#FFF2CC", f, c, fo, fi, False), 1, "pivot ignores colours"

    ' Canonical spelling counts as a match.
    f(1) = "$#,##0": c(1) = "\$#,##0"
    fo(1) = "": fi(1) = ""
    CheckEq Core_BestMatch("\$#,##0", "", "", f, c, fo, fi, True), 1, "canonical format"
End Sub

Private Sub Test_Colors()
    CheckEq Core_HexToColor("#FF0000"), RGB(255, 0, 0), "red"
    CheckEq Core_HexToColor("0000ff"), RGB(0, 0, 255), "blue no hash"
    CheckEq Core_HexToColor(""), -1, "blank"
    CheckEq Core_HexToColor("#12345"), -1, "too short"
    CheckEq Core_HexToColor("#GG0000"), -1, "not hex"
    CheckEq Core_ColorToHex(RGB(255, 242, 204)), "#FFF2CC", "to hex"
    CheckEq Core_ColorToHex(-1), "", "negative to hex"
    CheckEq Core_NormalizeHex(" #fff2cc "), "#FFF2CC", "normalize"
End Sub

Private Sub Test_Shortcuts()
    CheckEq Core_ShortcutToOnKey("Ctrl+Alt+1"), "^%1", "ctrl alt 1"
    CheckEq Core_ShortcutToOnKey("ctrl + shift + a"), "^+a", "ctrl shift a"
    CheckEq Core_ShortcutToOnKey("Alt+F5"), "%{F5}", "alt f5"
    CheckEq Core_ShortcutToOnKey("F9"), "{F9}", "bare function key"
    CheckEq Core_ShortcutToOnKey("Ctrl++"), "^{+}", "ctrl plus"
    CheckEq Core_ShortcutToOnKey("A"), "", "bare letter rejected"
    CheckEq Core_ShortcutToOnKey("Shift+A"), "", "shift only rejected"
    CheckEq Core_ShortcutToOnKey("Ctrl+Ctrl+A"), "", "repeated modifier"
    CheckEq Core_ShortcutToOnKey("Hyper+A"), "", "unknown modifier"
    CheckEq Core_ShortcutToOnKey(""), "", "blank"
    CheckEq Core_NormalizeShortcut("ctrl+alt+a"), "Ctrl+Alt+A", "normalize letter"
    CheckEq Core_NormalizeShortcut("alt+ctrl+shift+f12"), "Ctrl+Alt+Shift+F12", "normalize order"
    CheckEq Core_NormalizeShortcut("ctrl+pgdn"), "Ctrl+PgDn", "normalize named key"
    CheckEq Core_NormalizeShortcut("bogus"), "", "normalize invalid"
End Sub

Private Sub Test_XmlEscape()
    CheckEq Core_XmlEscape("A&B <""x"">"), "A&amp;B &lt;&quot;x&quot;&gt;", "xml escape"
    CheckEq Core_RibbonLabel("P&L"), "P&&L", "ribbon ampersand"
End Sub

Private Sub Test_Versions()
    CheckEq Core_CompareVersions("1.0.0", "1.0.0"), 0, "version equal"
    CheckEq Core_CompareVersions("v1.2.0", "1.1.9"), 1, "version newer, leading v"
    CheckEq Core_CompareVersions("1.9", "1.10"), -1, "version numeric not text"
    CheckEq Core_CompareVersions("1.1", "1.1.0"), 0, "version missing part is 0"
    CheckEq Core_CompareVersions("", "0.0.1"), -1, "version blank"
End Sub

' Digit's own icons are packed in the add-in, unpacked to the (test) presets
' folder and turned into pictures; Office icons are offered after them.
Private Sub Test_Icons()
    Dim c As Collection
    Check Icons_List().Count >= 36, "digit icons packed in the add-in"
    Check Icons_Exists("digit:pct_green"), "digit icon unpacked"
    Check Not Icons_Exists("digit:no_such_icon"), "unknown digit icon"
    Check Not Icons_Picture("digit:pct_green") Is Nothing, "digit icon picture for the ribbon"
    Check Not Icons_Picture("digit:pct_green", &HFFFFFF) Is Nothing, "digit icon picture for forms"
    Check UI_IconExists("digit:dec1_blue"), "UI knows digit icons"
    Check Not UI_IconExists("NoSuchOfficeIcon123"), "unknown office icon"
    Set c = UI_IconCandidates()
    Check c.Count > Icons_List().Count + 50, "picker offers office icons too"
    CheckEq Left$(c(1), 6), "digit:", "digit icons come first"
End Sub

' Loads the Settings window without showing it: runs its start-up code
' (presets, format list, icon grid) so errors there fail the build.
Private Sub Test_SettingsForm()
    Dim f As Object
    Set f = New frmDigitSettings
    Check f.Controls("lstGroups").ListCount > 0, "settings lists the preset groups"
    Check Not f.Controls("fraIcons").Visible, "icon picker starts closed"
    Unload f
End Sub

Private Sub Test_Collections()
    Dim c As New Collection
    c.Add "a": c.Add "b": c.Add "c"
    Coll_Swap c, 1, 3
    CheckEq c(1) & c(2) & c(3), "cba", "swap"
    Coll_ReplaceAt c, 2, "x"
    CheckEq c(1) & c(2) & c(3), "cxa", "replace"
    CheckEq c.Count, 3, "count kept"
End Sub

Private Sub Test_Normalize()
    Dim cfg As Scripting.Dictionary, json As String, errMsg As String
    json = "{""groups"":[" & _
           "{""name"":""A"",""presets"":[{""numberFormat"":""0""},{""numberFormat"":""0.0"",""fontColor"":""bad""}," & _
           "{""numberFormat"":""0.00"",""fillColor"":""#ffffff""},{""numberFormat"":""0.000""}]}," & _
           "{""name"":""a"",""presets"":[{""numberFormat"":""@""}]}," & _
           "{""name"":"""",""presets"":[{""numberFormat"":""@""}]}," & _
           "{""name"":""Empty"",""presets"":[]}" & _
           "],""slots"":[""A"",""missing"",null],""settings"":{""shortcuts"":[""ctrl+alt+q"",""nonsense""],""undoDepth"":500}}"
    Set cfg = Store_ParseConfig(json, errMsg)
    Check Not cfg Is Nothing, "parse ok"
    If cfg Is Nothing Then Exit Sub
    CheckEq cfg("groups").Count, 2, "nameless and empty groups dropped"
    CheckEq cfg("groups")(1)("presets").Count, 3, "max 3 presets"
    CheckEq cfg("groups")(1)("presets")(2)("fontColor"), "", "bad colour blanked"
    CheckEq cfg("groups")(1)("presets")(3)("fillColor"), "#FFFFFF", "colour upper-cased"
    CheckEq cfg("groups")(2)("name"), "a (2)", "duplicate name suffixed"
    CheckEq cfg("slots").Count, 5, "5 slots"
    CheckEq cfg("slots")(1), "A", "slot kept"
    CheckEq cfg("slots")(2), "", "unknown slot cleared"
    CheckEq cfg("slots")(3), "", "null slot cleared"
    CheckEq cfg("settings")("shortcuts")(1), "Ctrl+Alt+Q", "shortcut normalised"
    CheckEq cfg("settings")("shortcuts")(2), "", "bad shortcut cleared"
    CheckEq cfg("settings")("shortcuts").Count, 5, "5 shortcuts"
    CheckEq cfg("settings")("undoDepth"), 50, "undo depth clamped"
    CheckEq cfg("settings")("shortcutsEnabled"), True, "shortcuts default on"

    Check Store_ParseConfig("not json", errMsg) Is Nothing, "invalid json rejected"
    Check Store_ParseConfig("{""groups"":[]}", errMsg) Is Nothing, "no groups rejected"
End Sub

Private Sub Test_StoreRoundTrip()
    Dim cfg As Scripting.Dictionary, errMsg As String, g As Scripting.Dictionary

    Store_Load                                    ' no file yet -> starter presets written
    Check Len(Dir$(Store_FilePath())) > 0, "presets file created on first run"
    Set cfg = Store_Clone(Store_Config())
    CheckEq cfg("groups").Count, Store_DefaultConfig()("groups").Count, "starter groups loaded"

    Set g = Store_NewGroup("Unicode " & ChrW(&H20AC))
    g("presets").Add Store_NewPreset("#,##0 ""EUR""", "#123456")
    cfg("groups").Add g
    Store_RenameGroup cfg, "Numbers", "Plain numbers"
    CheckEq Store_SlotGroupName(cfg, 1), "Plain numbers", "rename updates slot"
    Store_DeleteGroup cfg, "Thousands"
    CheckEq Store_SlotGroupName(cfg, 2), "", "delete clears slot"
    Check Store_Save(cfg, errMsg), "save " & errMsg

    Store_Load
    Set cfg = Store_Config()
    Check Not Store_FindGroup(cfg, "Unicode " & ChrW(&H20AC)) Is Nothing, "unicode name survives save/load"
    CheckEq Store_FindGroup(cfg, "unicode " & ChrW(&H20AC))("presets")(1)("fontColor"), "#123456", "colour survives"
    Check Store_FindGroup(cfg, "Thousands") Is Nothing, "deleted group gone"
    Check Len(Dir$(Store_FilePath() & ".bak")) > 0, "backup written"
End Sub

Private Sub Test_FormatUtilities()
    Check Fmt_IsValidNumberFormat("#,##0.00"), "valid format"
    Check Not Fmt_IsValidNumberFormat(""), "blank invalid"
    Check Not Fmt_IsValidNumberFormat("0;0;0;0;0"), "five sections invalid"
    CheckEq Fmt_Preview("0.0%", 0.1234), Format$(0.1234, "0.0%"), "preview percent"
    CheckEq Fmt_Preview("#,##0;(#,##0);-", -1234), "(" & Format$(1234, "#,##0") & ")", "preview negative"
    CheckEq Fmt_Preview("#,##0;(#,##0);-", 0), "-", "preview zero"
End Sub

'------------------------------------------------------------------------------
' Integration tests
'------------------------------------------------------------------------------

Private Function TestGroup() As Scripting.Dictionary
    Dim g As Scripting.Dictionary
    Set g = Store_NewGroup("T")
    g("presets").Add Store_NewPreset("0.0%")
    g("presets").Add Store_NewPreset("#,##0.00")
    g("presets").Add Store_NewPreset("0.000")
    Set TestGroup = g
End Function

' Theme colour of a cell's font, or -1 when it is not a theme colour.
Private Function FontTheme(ByVal c As Range) As Variant
    FontTheme = -1
    On Error Resume Next
    FontTheme = c.Font.ThemeColor
End Function

Private Function NewScratchBook() As Workbook
    Set NewScratchBook = Workbooks.Add(xlWBATWorksheet)
    NewScratchBook.Worksheets(1).Activate
End Function

Private Function SameFormat(ByVal actual As Variant, ByVal fmt As String) As Boolean
    If IsNull(actual) Then Exit Function
    SameFormat = Core_FormatsEqual(CStr(actual), Fmt_Canonical(fmt))
End Function

Private Sub Test_RangeCycleAndUndo()
    Dim wb As Workbook, ws As Worksheet, g As Scripting.Dictionary, msg As String, res As DigitResult

    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Set g = TestGroup()
    Undo_Clear
    Undo_SetDepth 10

    ws.Range("A1:B3").Value = 0.5
    ws.Range("B2").NumberFormat = "@"           ' mixed formats in the selection
    ws.Range("A1:B3").Select
    ws.Range("A1").Activate

    res = Fmt_ApplyGroup(g, msg)
    CheckEq res, drOK, "apply 1: " & msg
    Check SameFormat(ws.Range("A1:B3").NumberFormat, "0.0%"), "no match -> preset 1 on whole selection"
    Fmt_ApplyGroup g, msg
    Check SameFormat(ws.Range("A1:B3").NumberFormat, "#,##0.00"), "preset 1 -> 2"
    Fmt_ApplyGroup g, msg
    Check SameFormat(ws.Range("A1:B3").NumberFormat, "0.000"), "preset 2 -> 3"
    Fmt_ApplyGroup g, msg
    Check SameFormat(ws.Range("A1:B3").NumberFormat, "0.0%"), "preset 3 loops to 1"
    CheckEq Undo_Count(), 4, "four undo entries"

    CheckEq Undo_Last(msg), drOK, "undo: " & msg
    Check SameFormat(ws.Range("A1").NumberFormat, "0.000"), "undo restores previous step"
    Undo_Last msg: Undo_Last msg: Undo_Last msg
    Check SameFormat(ws.Range("A1").NumberFormat, "General"), "full undo restores General"
    CheckEq ws.Range("B2").NumberFormat, "@", "full undo restores mixed cell"
    CheckEq Undo_Last(msg), drNothingToDo, "nothing left to undo"

    ' History depth
    Undo_SetDepth 2
    Fmt_ApplyGroup g, msg: Fmt_ApplyGroup g, msg: Fmt_ApplyGroup g, msg
    CheckEq Undo_Count(), 2, "history trimmed to depth"
    Undo_SetDepth 10

    ' Entire column: snapshot stays small and undo works.
    Undo_Clear
    ws.Range("C5").NumberFormat = "0.0"
    ws.Columns(3).Select
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "entire column: " & msg
    Check SameFormat(ws.Range("C1048576").NumberFormat, "0.0%"), "entire column formatted"
    Undo_Last msg
    CheckEq ws.Range("C5").NumberFormat, "0.0", "entire column undo (odd cell)"
    Check SameFormat(ws.Range("C100").NumberFormat, "General"), "entire column undo (rest)"

    ' Entire row, same idea across columns.
    Undo_Clear
    ws.Range("H7").NumberFormat = "0.0"
    ws.Rows(7).Select
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "entire row: " & msg
    Check SameFormat(ws.Range("XFD7").NumberFormat, "0.0%"), "entire row formatted"
    Undo_Last msg
    CheckEq ws.Range("H7").NumberFormat, "0.0", "entire row undo (odd cell)"
    Check SameFormat(ws.Range("Z7").NumberFormat, "General"), "entire row undo (rest)"

    ' Merged active cell
    ws.Range("E1:F2").Merge
    ws.Range("E1:F2").Select
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "merged: " & msg
    Check SameFormat(ws.Range("E1").NumberFormat, "0.0%"), "merged cell formatted"

    ' Shape selected: nothing happens.
    ws.Shapes.AddShape(1, 10, 10, 20, 20).Select
    CheckEq Fmt_ApplyGroup(g, msg), drNothingToDo, "shape selected is skipped"

    wb.Close SaveChanges:=False
End Sub

Private Sub Test_MultiAreaAndColours()
    Dim wb As Workbook, ws As Worksheet, g As Scripting.Dictionary, msg As String

    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Undo_Clear

    Set g = Store_NewGroup("Inputs")
    g("presets").Add Store_NewPreset("#,##0", "#0000FF")
    g("presets").Add Store_NewPreset("#,##0", "#0000FF", "#FFF2CC")

    Dim fontBefore As Long, themeBefore As Variant
    fontBefore = ws.Range("C1").Font.Color
    themeBefore = FontTheme(ws.Range("C1"))

    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    ws.Range("A1:A2,C1:C2").Select
    ws.Range("A1").Activate

    Fmt_ApplyGroup g, msg
    CheckEq ws.Range("C2").Font.Color, RGB(0, 0, 255), "multi-area: second area font"
    CheckEq ws.Range("A1").Interior.Color, RGB(255, 0, 0), "blank fill leaves fill alone"
    Fmt_ApplyGroup g, msg
    CheckEq ws.Range("C1").Interior.Color, RGB(255, 242, 204), "same format told apart by colour -> preset 2"
    CheckEq ws.Range("A1").Interior.Color, RGB(255, 242, 204), "fill applied to every area"

    Undo_Last msg: Undo_Last msg
    CheckEq ws.Range("A1").Interior.Color, RGB(255, 0, 0), "undo restores original fill"
    CheckEq ws.Range("C1").Interior.Pattern, xlNone, "undo restores no fill"
    CheckEq ws.Range("C1").Font.Color, fontBefore, "undo restores original font colour"
    CheckEq FontTheme(ws.Range("C1")), themeBefore, "undo restores theme font colour"

    wb.Close SaveChanges:=False
End Sub

Private Sub Test_TableColumn()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, g As Scripting.Dictionary, msg As String

    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Set g = TestGroup()
    ws.Range("A1:C1").Value = Array("Name", "Amount", "Rate")
    ws.Range("A2:C5").Value = 1
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C5"), , xlYes)

    ws.Range("B3").Select
    Fmt_ApplyGroup g, msg
    Check SameFormat(lo.ListColumns(2).DataBodyRange.NumberFormat, "0.0%"), "single cell formats whole table column"
    Check SameFormat(ws.Range("C2").NumberFormat, "General"), "other table columns untouched"
    Check SameFormat(ws.Range("B1").NumberFormat, "General"), "header untouched"

    ws.Range("C2:C3").Select
    Fmt_ApplyGroup g, msg
    Check SameFormat(ws.Range("C3").NumberFormat, "0.0%"), "multi-cell table selection formatted"
    Check SameFormat(ws.Range("C4").NumberFormat, "General"), "multi-cell table selection only"

    wb.Close SaveChanges:=False
End Sub

Private Sub Test_ProtectedSheet()
    Dim wb As Workbook, ws As Worksheet, g As Scripting.Dictionary, msg As String, res As DigitResult

    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Set g = TestGroup()
    Undo_Clear
    ws.Range("A1").Select
    ws.Protect
    res = Fmt_ApplyGroup(g, msg)
    CheckEq res, drWarning, "protected sheet gives a friendly warning"
    Check InStr(1, msg, "protected", vbTextCompare) > 0, "message mentions protection"
    CheckEq Undo_Count(), 0, "nothing pushed to undo"
    ws.Unprotect
    wb.Close SaveChanges:=False
End Sub

Private Sub Test_PivotDataField()
    Dim wb As Workbook, ws As Worksheet, pc As PivotCache, pt As PivotTable, g As Scripting.Dictionary
    Dim msg As String, df As PivotField, valueCell As Range

    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Set g = TestGroup()
    Undo_Clear

    ws.Range("A1:B1").Value = Array("Region", "Sales")
    ws.Range("A2:B2").Value = Array("North", 100)
    ws.Range("A3:B3").Value = Array("South", 250)
    ws.Range("A4:B4").Value = Array("North", 50)

    Set pc = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B4"))
    Set pt = pc.CreatePivotTable(ws.Range("E3"), "DigitTestPivot")
    pt.PivotFields("Region").Orientation = xlRowField
    Set df = pt.AddDataField(pt.PivotFields("Sales"), "Total Sales", xlSum)

    Set valueCell = df.DataRange.Cells(1, 1)
    valueCell.Select
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "pivot apply: " & msg
    Check SameFormat(pt.DataFields("Total Sales").NumberFormat, "0.0%"), "pivot DataField number format set"
    pt.RefreshTable
    Check SameFormat(pt.DataFields("Total Sales").NumberFormat, "0.0%"), "format survives refresh"

    valueCell.Select
    Fmt_ApplyGroup g, msg
    Check SameFormat(pt.DataFields("Total Sales").NumberFormat, "#,##0.00"), "pivot cycles"
    Undo_Last msg
    Check SameFormat(pt.DataFields("Total Sales").NumberFormat, "0.0%"), "pivot undo"

    wb.Close SaveChanges:=False
End Sub
