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
    Test_Classify
    Test_Normalize
    Test_ContentSettings
    Test_StoreRoundTrip
    Test_FormatUtilities
    Test_RangeCycleAndUndo
    Test_MultiAreaAndColours
    Test_Styles
    Test_ColourByContent
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
    Check Icons_Exists("digit:colour"), "colour by content icon unpacked"
    Check Not Icons_Exists("digit:no_such_icon"), "unknown digit icon"
    Check Not Icons_Picture("digit:pct_green") Is Nothing, "digit icon picture for the ribbon"
    Check Not Icons_Picture("digit:pct_green", &HFFFFFF) Is Nothing, "digit icon picture for forms"
    Check UI_IconExists("digit:dec1_blue"), "UI knows digit icons"
    Check Not Icons_FormPicture("digit:pct_green") Is Nothing, "digit icon loads in the settings window"
    Check Icons_FormPicture("digit:no_such_icon") Is Nothing, "unknown digit icon has no picture"
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
    CheckEq f.Controls("mpTabs").Pages.Count, 3, "settings has three tabs"
    CheckEq f.Controls("txtKeyColour").Text, Store_ContentShortcut(Store_Config()), "colour by content shortcut shown"
    Check Len(f.Controls("lblCcSwatch4").ControlTipText) > 0, "colour by content swatches filled"
    CheckEq f.Controls("cboCcSlot").ListCount, 6, "colour by content position choices"
    CheckEq f.Controls("cboCcSlot").ListIndex, Store_ColourSlot(Store_Config()), "colour by content position shown"
    Check f.Controls("chkBold").TripleState, "bold box has a leave-as-is state"
    CheckEq Left$(f.Controls("lblSlot1").Caption, 1), "1", "shortcut rows are numbered"
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

' Colour by content: the formula classifier. The cell is on sheet "Model".
Private Sub Test_Classify()
    Dim names As New Scripting.Dictionary, sheets As New Scripting.Dictionary, v As Variant

    For Each v In Array("model", "inputs", "my sheet", "jan", "dec")
        sheets(CStr(v)) = True
    Next v
    names("taxrate") = "=0.05"
    names("revenue") = "=Inputs!$C$7"
    names("local") = "=Model!$A$1"
    names("ext") = "='[Budget.xlsx]Sheet1'!$A$1"
    names("chain") = "=Revenue*2"
    names("model!scoped") = "=Inputs!$A$1"
    names("scoped") = "=Model!$A$1"
    names("table1") = "='Inputs'!A1"
    names("table2") = "='Model'!A1"
    names("loop") = "=Loop+1"

    CheckEq Core_ClassifyFormula("1250", "Model", names, sheets), ckNone, "classify: not a formula"

    For Each v In Array("=1000*12", "=-(5+3)/2", "=1E+3*2", "=TRUE", "=.5*2", "={1,2;3,4}")
        CheckEq Core_ClassifyFormula(CStr(v), "Model", names, sheets), ckInput, "classify input: " & v
    Next v

    For Each v In Array("=B5*C5", "=SUM(D2:D10)", "=TODAY()", "=B5*1.05", "=$A$1", "=SUM(A:A)", "=SUM(1:1)", _
                        "=LOG10(100)", "=Model!A1", "=""See Inputs!A1""", "=""a""&""b""", "=A1&"" units""", _
                        "=#REF!+1", "=#N/A", "=#DIV/0!", "=TaxRate*B5", "=Local", "=Table2[Amount]", _
                        "=[@Amount]*2", "=Table2[[#This Row],[Amount]]", "=Table2[[#This Row],[It's '[x']]]]", _
                        "=LET(x,1,x*2)", "=Undefined+1", "=RC[-1]*R[-1]C", "=A1#", "=Loop", "=Model!1:1+TaxRate")
        CheckEq Core_ClassifyFormula(CStr(v), "Model", names, sheets), ckFormula, "classify formula: " & v
    Next v

    For Each v In Array("=Inputs!C7", "='My Sheet'!A1", "=B5*Inputs!C2", "=SUM(Jan:Dec!B5)", "=SUM('Jan:Dec'!B5)", _
                        "=Revenue", "=Chain", "=Scoped", "=Table1[Amount]", "=SUM(Table1[[#Totals],[Amount]])", _
                        "=Inputs!RC[-1]", "=Model!1:1+Revenue", "=#REF!A1+Inputs!A1")
        CheckEq Core_ClassifyFormula(CStr(v), "Model", names, sheets), ckSheetLink, "classify sheet link: " & v
    Next v

    For Each v In Array("=[Budget.xlsx]Sheet1!A1", "=[1]Sheet1!A1", "='C:\path\[Budget.xlsx]Sheet1'!A1", _
                        "=Budget.xlsx!Name", "='C:\x\Budget.xlsx'!Name", "=PERSONAL.XLSB!MyFn(A1)", _
                        "=Inputs!A1+[Budget.xlsx]Sheet1!A1", "=Ext", "=[Budget.xlsx]Jan:Dec!A1", _
                        "='\\srv\share\Budget.xlsx'!Name")
        CheckEq Core_ClassifyFormula(CStr(v), "Model", names, sheets), ckExternal, "classify external: " & v
    Next v

    ' Without the sheet list or names, any unbracketed qualifier is a sheet.
    CheckEq Core_ClassifyFormula("=Assumptions!C7", "Model", Nothing, Nothing), ckSheetLink, "classify: no sheet list"
    CheckEq Core_ClassifyFormula("=Model!C7", "model", Nothing, Nothing), ckFormula, "classify: own sheet, any case"
    CheckEq Core_ClassifyFormula("=Scoped", "Inputs", names, sheets), ckSheetLink, "classify: global name from another sheet"
    sheets("bob's p&l") = True
    CheckEq Core_ClassifyFormula("='Bob''s P&L'!A1", "Model", names, sheets), ckSheetLink, "classify: apostrophe in a sheet name"
    sheets("model.xlsx") = "workbook"
    CheckEq Core_ClassifyFormula("=Model.xlsx!TaxRate", "Model", names, sheets), ckFormula, "classify: this workbook's own name"
    CheckEq Core_ClassifyFormula("=Other.xlsx!TaxRate", "Model", names, sheets), ckExternal, "classify: another workbook's name"
End Sub

Private Sub Test_Normalize()
    Dim cfg As Scripting.Dictionary, json As String, errMsg As String
    json = "{""groups"":[" & _
           "{""name"":""A"",""presets"":[{""numberFormat"":""0""},{""numberFormat"":""0.0"",""fontColor"":""bad""}," & _
           "{""numberFormat"":""0.00"",""fillColor"":""#ffffff""},{""numberFormat"":""0.000""}]}," & _
           "{""name"":""a"",""presets"":[{""numberFormat"":""@""}]}," & _
           "{""name"":"""",""presets"":[{""numberFormat"":""@""}]}," & _
           "{""name"":""Empty"",""presets"":[]}" & _
           ",{""name"":""B"",""presets"":[{""numberFormat"":""0""}]}" & _
           "],""slots"":[""B"",""missing"",null],""settings"":{""shortcuts"":[""ctrl+alt+q"",""nonsense""],""undoDepth"":500}}"
    Set cfg = Store_ParseConfig(json, errMsg)
    Check Not cfg Is Nothing, "parse ok"
    If cfg Is Nothing Then Exit Sub
    CheckEq cfg("groups").Count, 3, "nameless and empty groups dropped"
    CheckEq cfg("groups")(1)("name"), "B", "group on a button in an old file moves to the front"
    CheckEq cfg("groups")(2)("presets").Count, 3, "max 3 presets"
    CheckEq cfg("groups")(2)("presets")(2)("fontColor"), "", "bad colour blanked"
    CheckEq cfg("groups")(2)("presets")(3)("fillColor"), "#FFFFFF", "colour upper-cased"
    CheckEq cfg("groups")(3)("name"), "a (2)", "duplicate name suffixed"
    CheckEq cfg("slots").Count, 5, "5 slots"
    CheckEq cfg("slots")(1), "B", "slots follow the group order"
    CheckEq cfg("slots")(2), "A", "slot 2 is the second group"
    CheckEq cfg("slots")(4), "", "no group for slot 4"
    CheckEq Store_SlotGroupName(cfg, 3), "a (2)", "quick button 3 is the third group"
    Check Not Store_IsInSlot(cfg, "missing"), "unknown group is not on a button"
    CheckEq cfg("settings")("shortcuts")(1), "Ctrl+Alt+Q", "shortcut normalised"
    CheckEq cfg("settings")("shortcuts")(2), "", "bad shortcut cleared"
    CheckEq cfg("settings")("shortcuts").Count, 5, "5 shortcuts"
    CheckEq cfg("settings")("undoDepth"), 50, "undo depth clamped"
    CheckEq cfg("settings")("shortcutsEnabled"), True, "shortcuts default on"
    CheckEq Store_LargeButtons(cfg), False, "small quick buttons by default"
    Set cfg = Store_ParseConfig("{""groups"":[{""name"":""X"",""presets"":[{""numberFormat"":""0""}]}],""settings"":{""largeButtons"":true}}", errMsg)
    CheckEq Store_LargeButtons(cfg), True, "large quick buttons setting read"
    Set cfg = Store_ParseConfig(Store_ToJson(cfg), errMsg)
    CheckEq Store_LargeButtons(cfg), True, "large quick buttons setting survives save"

    Check Store_ParseConfig("not json", errMsg) Is Nothing, "invalid json rejected"
    Check Store_ParseConfig("{""groups"":[]}", errMsg) Is Nothing, "no groups rejected"
End Sub

' Colour by content settings: defaults, repair, and the shortcut giving way.
Private Sub Test_ContentSettings()
    Dim cfg As Scripting.Dictionary, errMsg As String, grp As String
    grp = """groups"":[{""name"":""X"",""presets"":[{""numberFormat"":""0""}]}]"

    Set cfg = Store_ParseConfig("{" & grp & "}", errMsg)
    Check Not cfg Is Nothing, "content settings: parse ok"
    If cfg Is Nothing Then Exit Sub
    CheckEq Store_ContentColor(cfg, ckInput), "#0000FF", "default input colour"
    CheckEq Store_ContentColor(cfg, ckFormula), "#000000", "default formula colour"
    CheckEq Store_ContentColor(cfg, ckSheetLink), "#008000", "default sheet link colour"
    CheckEq Store_ContentColor(cfg, ckExternal), "#C00000", "default external colour"
    CheckEq Store_ContentShortcut(cfg), "Ctrl+Alt+6", "default content shortcut"

    Set cfg = Store_ParseConfig("{" & grp & ",""settings"":{""contentColors"":" & _
                                "{""input"":"""",""formula"":""zzz"",""sheetLink"":""#00ff00""}," & _
                                """contentColorShortcut"":""ctrl+shift+k""}}", errMsg)
    CheckEq Store_ContentColor(cfg, ckInput), "", "blank colour kept (skip this kind)"
    CheckEq Store_ContentColor(cfg, ckFormula), "#000000", "invalid colour repaired to the default"
    CheckEq Store_ContentColor(cfg, ckSheetLink), "#00FF00", "colour upper-cased"
    CheckEq Store_ContentColor(cfg, ckExternal), "#C00000", "missing colour gets the default"
    CheckEq Store_ContentShortcut(cfg), "Ctrl+Shift+K", "content shortcut normalised"

    Set cfg = Store_Clone(cfg)
    CheckEq Store_ContentColor(cfg, ckInput), "", "blank colour survives a save"
    CheckEq Store_ContentColor(cfg, ckSheetLink), "#00FF00", "colour survives a save"
    CheckEq Store_ContentShortcut(cfg), "Ctrl+Shift+K", "content shortcut survives a save"

    Store_SetContentColor cfg, ckExternal, "#ff0000"
    CheckEq Store_ContentColor(cfg, ckExternal), "#FF0000", "set colour"
    Store_SetContentColor cfg, ckExternal, ""
    CheckEq Store_ContentColor(cfg, ckExternal), "", "clear colour"

    ' A quick button that already uses Ctrl+Alt+6 keeps it.
    Set cfg = Store_ParseConfig("{" & grp & ",""settings"":{""shortcuts"":[""Ctrl+Alt+6"","""","""","""",""""]}}", errMsg)
    CheckEq Store_ContentShortcut(cfg), "", "content shortcut gives way to a quick button"
    CheckEq cfg("settings")("shortcuts")(1), "Ctrl+Alt+6", "quick button keeps its shortcut"

    ' Colour by content in place of quick button 4: the groups after it move along.
    grp = """groups"":[{""name"":""A"",""presets"":[{""numberFormat"":""0""}]},{""name"":""B"",""presets"":[{""numberFormat"":""0""}]}," & _
          "{""name"":""C"",""presets"":[{""numberFormat"":""0""}]},{""name"":""D"",""presets"":[{""numberFormat"":""0""}]}," & _
          "{""name"":""E"",""presets"":[{""numberFormat"":""0""}]},{""name"":""F"",""presets"":[{""numberFormat"":""0""}]}]"
    Set cfg = Store_ParseConfig("{" & grp & "}", errMsg)
    CheckEq Store_ColourSlot(cfg), 0, "colour by content has its own button by default"
    CheckEq Store_SlotGroupName(cfg, 5), "E", "own button: five groups on quick buttons"
    CheckEq Store_GroupSlot(cfg, 5), 5, "own button: group 5 is on button 5"
    Set cfg = Store_ParseConfig("{" & grp & ",""settings"":{""contentColorSlot"":4}}", errMsg)
    CheckEq Store_ColourSlot(cfg), 4, "colour slot read"
    CheckEq Store_QuickGroupCount(cfg), 4, "colour slot: four groups on quick buttons"
    CheckEq Store_SlotGroupName(cfg, 3), "C", "colour slot: positions before it unchanged"
    CheckEq Store_SlotGroupName(cfg, 4), "", "colour slot: no group at its position"
    CheckEq Store_SlotGroupName(cfg, 5), "D", "colour slot: next group moves along"
    CheckEq Store_GroupSlot(cfg, 3), 3, "colour slot: group 3 on button 3"
    CheckEq Store_GroupSlot(cfg, 4), 5, "colour slot: group 4 on button 5"
    CheckEq Store_GroupSlot(cfg, 5), 0, "colour slot: group 5 in the menu"
    Check Store_IsInSlot(cfg, "D"), "colour slot: group 4 is on a button"
    Check Not Store_IsInSlot(cfg, "E"), "colour slot: group 5 is not on a button"
    Set cfg = Store_Clone(cfg)
    CheckEq Store_ColourSlot(cfg), 4, "colour slot survives a save"
    CheckEq cfg("groups")(4)("name") & cfg("groups")(5)("name"), "DE", "colour slot: group order survives a save"
    Store_SetColourSlot cfg, 1
    CheckEq Store_SlotGroupName(cfg, 2), "A", "colour slot 1: first group on button 2"
    Set cfg = Store_ParseConfig("{" & grp & ",""settings"":{""contentColorSlot"":9}}", errMsg)
    CheckEq Store_ColourSlot(cfg), 0, "invalid colour slot reset"
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
    CheckEq Store_SlotGroupName(cfg, 1), "Plain numbers", "rename keeps the button"
    Store_DeleteGroup cfg, "Thousands"
    CheckEq Store_SlotGroupName(cfg, 2), "Percent", "delete moves the next group up"
    Coll_Swap cfg("groups"), 1, 2
    CheckEq Store_SlotGroupName(cfg, 1), "Percent", "moving a group moves its button"
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

' Bold / italic / underline: stored per preset, used to tell presets apart,
' applied, and undone.
Private Sub Test_Styles()
    Dim wb As Workbook, ws As Worksheet, g As Scripting.Dictionary, msg As String, cfg As Scripting.Dictionary
    Dim errMsg As String, p As Scripting.Dictionary

    ' Store
    Set p = Store_NewPreset("0.0", "", "", "10-")
    CheckEq Store_PresetStyle(p), "10-", "style kept"
    Check Not p.Exists("underline"), "leave-as-is style is not stored"
    Store_SetPresetStyle p, "underline", True
    Store_SetPresetStyle p, "bold", Null
    CheckEq Store_PresetStyle(p), "-01", "style flags set and cleared"
    Set cfg = Store_ParseConfig("{""groups"":[{""name"":""S"",""presets"":[{""numberFormat"":""0"",""bold"":true,""italic"":false,""underline"":""x""}]}]}", errMsg)
    CheckEq Store_PresetStyle(cfg("groups")(1)("presets")(1)), "10-", "style read from the file (non-boolean ignored)"
    Set cfg = Store_ParseConfig(Store_ToJson(cfg), errMsg)
    CheckEq Store_PresetStyle(cfg("groups")(1)("presets")(1)), "10-", "style survives save"

    ' Matching: same number format, told apart by style
    CheckEq Core_BestMatch("0.0", "", "", Array("", "0.0", "0.0"), Array("", "0.0", "0.0"), Array("", "", ""), Array("", "", ""), _
                           True, "100", Array("", "1--", "1-1")), 1, "bold cell matches the bold preset"
    CheckEq Core_BestMatch("0.0", "", "", Array("", "0.0", "0.0"), Array("", "0.0", "0.0"), Array("", "", ""), Array("", "", ""), _
                           True, "101", Array("", "1--", "1-1")), 2, "bold underlined cell matches the more specific preset"
    CheckEq Core_BestMatch("0.0", "", "", Array("", "0.0"), Array("", "0.0"), Array("", ""), Array("", ""), _
                           True, "000", Array("", "1--")), 0, "plain cell does not match a bold preset"

    ' Apply and undo
    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    Undo_Clear
    Set g = Store_NewGroup("Styles")
    g("presets").Add Store_NewPreset("0.0", "", "", "1--")
    g("presets").Add Store_NewPreset("0.0", "", "", "1-1")
    ws.Range("A1:B2").Value = 1.5
    ws.Range("A1:B2").Select
    ws.Range("A1").Activate
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "style 1: " & msg
    CheckEq ws.Range("B2").Font.Bold, True, "bold applied"
    CheckEq ws.Range("B2").Font.Underline, xlUnderlineStyleNone, "underline left alone"
    CheckEq Fmt_ApplyGroup(g, msg), drOK, "style 2: " & msg
    CheckEq ws.Range("B2").Font.Underline, xlUnderlineStyleSingle, "underline applied by the second preset"
    Undo_Last msg
    CheckEq ws.Range("B2").Font.Underline, xlUnderlineStyleNone, "undo removes underline"
    CheckEq ws.Range("B2").Font.Bold, True, "undo keeps the earlier bold"
    Undo_Last msg
    CheckEq ws.Range("B2").Font.Bold, False, "undo removes bold"
    wb.Close SaveChanges:=False
End Sub

' Colour by content on a small model: every kind, nothing but the font colour
' changes, undo, and the selections that need care.
Private Sub Test_ColourByContent()
    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet, wsIn As Worksheet
    Dim msg As String, res As DigitResult, themeBefore As Variant
    Dim blue As Long, black As Long, green As Long, red As Long, pink As Long, yellow As Long

    blue = RGB(0, 0, 255): black = RGB(0, 0, 0): green = RGB(0, 128, 0): red = RGB(192, 0, 0)
    pink = RGB(255, 0, 255): yellow = RGB(255, 242, 204)

    Set wb2 = NewScratchBook()
    wb2.Worksheets(1).Range("A1").Value = 7
    Set wb = NewScratchBook()
    Set ws = wb.Worksheets(1)
    ws.Name = "Model"
    Set wsIn = wb.Worksheets.Add(After:=ws)
    wsIn.Name = "Inputs"
    wsIn.Range("C7").Value = 100
    wb.Names.Add Name:="Revenue", RefersTo:="=Inputs!$C$7"
    ws.Activate
    Undo_Clear
    Undo_SetDepth 10

    With ws
        .Range("A1").Value = 0.05
        .Range("A1").NumberFormat = "0.0%"
        .Range("A1").Font.Bold = True
        .Range("A1").Interior.Color = yellow
        .Range("A2").Formula = "=A1*2"
        .Range("A3").Formula = "=Inputs!C7*A1"
        .Range("A4").Formula = "=Revenue"
        .Range("A5").Formula = "=1000*12"
        .Range("A6").Value = True
        .Range("A7").Value = "Label"
        .Range("A7").Font.Color = pink
        .Range("A9").Formula = "=""See Inputs!C7"""
        .Range("B1").Value = DateSerial(2026, 1, 31)
        .Range("B2").Formula = "=SUM(A1:A2)"
        .Range("B3").Formula = "=[" & wb2.Name & "]" & wb2.Worksheets(1).Name & "!$A$1"
        .Range("B8").Font.Color = pink              ' blank cell
    End With
    themeBefore = FontTheme(ws.Range("A2"))

    ws.Range("A1:B10").Select
    res = Fmt_ColourByContent(msg)
    CheckEq res, drOK, "colour by content: " & msg
    CheckEq msg, "Digit: Colour by content  4 input, 3 formula, 2 sheet link, 1 external", "colour summary"
    CheckEq ws.Range("A1").Font.Color, blue, "typed number is an input"
    CheckEq ws.Range("B1").Font.Color, blue, "typed date is an input"
    CheckEq ws.Range("A6").Font.Color, blue, "typed TRUE is an input"
    CheckEq ws.Range("A5").Font.Color, blue, "formula of numbers only is an input"
    CheckEq ws.Range("A2").Font.Color, black, "same-sheet formula"
    CheckEq ws.Range("B2").Font.Color, black, "same-sheet function"
    CheckEq ws.Range("A9").Font.Color, black, "sheet name inside text is not a link"
    CheckEq ws.Range("A3").Font.Color, green, "link to another sheet"
    CheckEq ws.Range("A4").Font.Color, green, "name that points at another sheet"
    CheckEq ws.Range("B3").Font.Color, red, "link to another workbook"
    CheckEq ws.Range("A7").Font.Color, pink, "typed text is left alone"
    CheckEq ws.Range("B8").Font.Color, pink, "blank cell is left alone"
    Check SameFormat(ws.Range("A1").NumberFormat, "0.0%"), "number format untouched"
    CheckEq ws.Range("A1").Font.Bold, True, "bold untouched"
    CheckEq ws.Range("A1").Interior.Color, yellow, "fill untouched"
    CheckEq Undo_Count(), 1, "one undo entry"
    CheckEq Undo_PeekLabel(), "Colour by content", "undo label"

    ' Undo puts the font colours back and nothing else.
    ws.Range("A1").NumberFormat = "0.00"
    CheckEq Undo_Last(msg), drOK, "undo colour by content: " & msg
    Check ws.Range("A3").Font.Color <> green, "undo restores the font colour"
    CheckEq FontTheme(ws.Range("A2")), themeBefore, "undo restores a theme colour as a theme colour"
    Check SameFormat(ws.Range("A1").NumberFormat, "0.00"), "undo leaves a later number format alone"
    CheckEq ws.Range("A7").Font.Color, pink, "undo leaves text alone"

    ' One cell selected: only that cell.
    ws.Range("A1").Select
    CheckEq Fmt_ColourByContent(msg), drOK, "single cell: " & msg
    CheckEq ws.Range("A1").Font.Color, blue, "single cell coloured"
    Check ws.Range("A3").Font.Color <> green, "single cell: other cells untouched"
    ws.Range("A3").Select
    Fmt_ColourByContent msg
    CheckEq ws.Range("A3").Font.Color, green, "single formula cell coloured"
    Check ws.Range("B3").Font.Color <> red, "single formula cell: other cells untouched"

    ' Nothing but text and blanks: nothing happens.
    Undo_Clear
    ws.Range("A7:A8").Select
    CheckEq Fmt_ColourByContent(msg), drNothingToDo, "text only is skipped"
    CheckEq Undo_Count(), 0, "text only: no undo entry"
    ws.Range("A7").Select
    CheckEq Fmt_ColourByContent(msg), drNothingToDo, "single text cell is skipped"
    ws.Range("D20").Select
    CheckEq Fmt_ColourByContent(msg), drNothingToDo, "cell outside the used range is skipped"

    ' A blank colour in Settings leaves that kind alone.
    Store_SetContentColor Store_Config(), ckInput, ""
    ws.Range("A1").Font.Color = pink
    ws.Range("A2").Font.Color = pink
    ws.Range("A1:A2").Select
    CheckEq Fmt_ColourByContent(msg), drOK, "blank input colour: " & msg
    CheckEq ws.Range("A1").Font.Color, pink, "blank colour: inputs left alone"
    CheckEq ws.Range("A2").Font.Color, black, "blank colour: other kinds still coloured"
    Store_SetContentColor Store_Config(), ckInput, Store_ContentDefault(ckInput)

    ' A filled-down block with a different formula in the middle.
    ws.Range("D1:D50").Formula = "=A1+1"
    ws.Range("D25").Formula = "=Inputs!C7"
    ws.Range("E1:E50").Value = 3
    ws.Range("D1:E50").Select
    CheckEq Fmt_ColourByContent(msg), drOK, "block: " & msg
    CheckEq ws.Range("D24").Font.Color, black, "block: above the odd cell"
    CheckEq ws.Range("D25").Font.Color, green, "block: the odd cell"
    CheckEq ws.Range("D50").Font.Color, black, "block: below the odd cell"
    CheckEq ws.Range("E50").Font.Color, blue, "block: constants"

    ' Whole column and whole sheet stay within the used range.
    ws.Columns(1).Select
    CheckEq Fmt_ColourByContent(msg), drOK, "entire column: " & msg
    CheckEq ws.Range("A4").Font.Color, green, "entire column coloured"
    ws.Cells.Select
    CheckEq Fmt_ColourByContent(msg), drOK, "entire sheet: " & msg

    ' Several areas.
    ws.Range("A1:B3").Font.Color = pink
    ws.Range("A1,B3").Select
    Fmt_ColourByContent msg
    CheckEq ws.Range("A1").Font.Color, blue, "multi-area: first area"
    CheckEq ws.Range("B3").Font.Color, red, "multi-area: second area"
    CheckEq ws.Range("A2").Font.Color, pink, "multi-area: cells between are untouched"

    ' One merged cell selected: only that cell.
    ws.Range("G1").Value = 5
    ws.Range("G1:H1").Merge
    ws.Range("A1").Font.Color = pink
    ws.Range("G1:H1").Select
    CheckEq Fmt_ColourByContent(msg), drOK, "merged cell: " & msg
    CheckEq ws.Range("G1").Font.Color, blue, "merged cell coloured"
    CheckEq ws.Range("A1").Font.Color, pink, "merged cell: other cells untouched"

    ' Shape selected, then a protected sheet.
    ws.Shapes.AddShape(1, 10, 10, 20, 20).Select
    CheckEq Fmt_ColourByContent(msg), drNothingToDo, "shape selected is skipped"
    ws.Range("A1:B10").Select
    ws.Protect
    CheckEq Fmt_ColourByContent(msg), drWarning, "protected sheet gives a warning"
    ws.Unprotect

    wb.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
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
