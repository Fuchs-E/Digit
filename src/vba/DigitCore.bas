Attribute VB_Name = "DigitCore"
'==============================================================================
' DigitCore - pure logic shared by every layer.
'
' Nothing in this module touches Excel objects (no Range, Workbook, ribbon or
' UserForm). It holds the cycle rules, colour and shortcut conversion and small
' collection helpers, so it can be unit-tested on its own and translated line
' for line to C# in the COM add-in phase.
'==============================================================================
Option Explicit

Public Const DIGIT_SLOT_COUNT As Long = 5
Public Const DIGIT_MAX_PRESETS As Long = 3

' Result of a formatting or undo request, returned by the engine to the UI layer.
Public Enum DigitResult
    drOK = 0            ' done; show a status-bar note
    drNothingToDo = 1   ' nothing changed; status-bar note only
    drWarning = 2       ' nothing changed; tell the user in a message box
    drUndoTooBig = 3    ' too large to snapshot; ask before applying without undo
    drError = 4         ' unexpected failure; message box
End Enum

' What a cell contains, for Colour by content. The order is the precedence: a
' formula that mixes kinds takes the highest.
Public Enum DigitContentKind
    ckNone = 0          ' not a formula / leave alone
    ckInput = 1         ' typed number, or a formula of numbers and operators only
    ckFormula = 2       ' formula that stays on its own sheet
    ckSheetLink = 3     ' formula that points at another sheet in the workbook
    ckExternal = 4      ' formula that points at another workbook
End Enum

Public Const DIGIT_CONTENT_KINDS As Long = 4

'------------------------------------------------------------------------------
' Cycle logic
'------------------------------------------------------------------------------

' Returns the 1-based preset to apply next, given the 1-based preset the target
' currently matches (0 = matches none) and the number of presets in the group.
'   matches none  -> 1
'   matches k     -> k + 1
'   matches last  -> 1 (loop)
Public Function Core_NextIndex(ByVal matchIndex As Long, ByVal presetCount As Long) As Long
    If presetCount < 1 Then
        Core_NextIndex = 0
    ElseIf matchIndex < 1 Or matchIndex >= presetCount Then
        Core_NextIndex = 1
    Else
        Core_NextIndex = matchIndex + 1
    End If
End Function

' Finds which preset the current cell state matches. Arrays are 1-based and the
' same length. presetCanon holds the format string as Excel stores it after
' applying it (Excel sometimes rewrites formats, e.g. $ becomes \$), so both
' spellings count as a match.
'
' When useColours is True, a preset that also specifies a font or fill colour
' must match that colour too, and a more specific match wins over a less specific
' one. This keeps groups such as "blue font" / "blue font + yellow fill" cycling
' correctly even when they share a number format. Blank colours are wildcards
' because a blank colour means "leave as is".
'
' Bold, italic and underline work the same way (when useColours is True):
' curStyle is the cell's state ("1" on, "0" off, "?" mixed) for bold, italic,
' underline, and each preset style uses "1" / "0", or "-" for leave as is.
'
' Returns 0 when nothing matches.
Public Function Core_BestMatch(ByVal curFormat As String, ByVal curFont As String, ByVal curFill As String, _
                               presetFormats As Variant, presetCanon As Variant, _
                               presetFonts As Variant, presetFills As Variant, _
                               ByVal useColours As Boolean, _
                               Optional ByVal curStyle As String = "", Optional presetStyles As Variant) As Long
    Dim i As Long, score As Long, bestScore As Long, best As Long, k As Long, want As String

    For i = LBound(presetFormats) To UBound(presetFormats)
        If Core_FormatsEqual(curFormat, CStr(presetFormats(i))) Or _
           Core_FormatsEqual(curFormat, CStr(presetCanon(i))) Then
            score = 1
            If useColours Then
                If Len(presetFonts(i)) > 0 Then
                    If StrComp(CStr(presetFonts(i)), curFont, vbTextCompare) = 0 Then
                        score = score + 1
                    Else
                        score = -1000
                    End If
                End If
                If Len(presetFills(i)) > 0 Then
                    If StrComp(CStr(presetFills(i)), curFill, vbTextCompare) = 0 Then
                        score = score + 1
                    Else
                        score = -1000
                    End If
                End If
                If Not IsMissing(presetStyles) Then
                    For k = 1 To 3
                        want = Mid$(CStr(presetStyles(i)) & "---", k, 1)
                        If want = "0" Or want = "1" Then
                            If Mid$(curStyle & "???", k, 1) = want Then score = score + 1 Else score = -1000
                        End If
                    Next k
                End If
            End If
            If score > bestScore Then
                bestScore = score
                best = i
            End If
        End If
    Next i

    Core_BestMatch = best
End Function

' Number format strings are compared case-insensitively ("MMM-YY" = "mmm-yy").
Public Function Core_FormatsEqual(ByVal a As String, ByVal b As String) As Boolean
    If Len(a) = 0 Or Len(b) = 0 Then
        Core_FormatsEqual = False
    Else
        Core_FormatsEqual = (StrComp(a, b, vbTextCompare) = 0)
    End If
End Function

'------------------------------------------------------------------------------
' Colours. The presets file stores "#RRGGBB"; Excel and MSForms use BGR Longs.
'------------------------------------------------------------------------------

' "#RRGGBB" or "RRGGBB" -> Excel colour Long. Returns -1 for blank or invalid.
Public Function Core_HexToColor(ByVal hexText As String) As Long
    Dim s As String
    s = Trim$(hexText)
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Not Core_IsHex6(s) Then
        Core_HexToColor = -1
        Exit Function
    End If
    Core_HexToColor = RGB(CLng("&H" & Mid$(s, 1, 2)), CLng("&H" & Mid$(s, 3, 2)), CLng("&H" & Mid$(s, 5, 2)))
End Function

' Excel colour Long -> "#RRGGBB". Negative values return "".
Public Function Core_ColorToHex(ByVal colorValue As Long) As String
    If colorValue < 0 Then
        Core_ColorToHex = ""
        Exit Function
    End If
    Core_ColorToHex = "#" & Hex2(colorValue And &HFF&) & _
                            Hex2((colorValue \ &H100&) And &HFF&) & _
                            Hex2((colorValue \ &H10000) And &HFF&)
End Function

' Cleans a colour from the presets file: returns "#RRGGBB" in upper case, or ""
' when blank or invalid.
Public Function Core_NormalizeHex(ByVal hexText As String) As String
    Dim c As Long
    c = Core_HexToColor(hexText)
    If c < 0 Then
        Core_NormalizeHex = ""
    Else
        Core_NormalizeHex = Core_ColorToHex(c)
    End If
End Function

Private Function Core_IsHex6(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) <> 6 Then Exit Function
    For i = 1 To 6
        If InStr(1, "0123456789ABCDEFabcdef", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    Core_IsHex6 = True
End Function

Private Function Hex2(ByVal v As Long) As String
    Hex2 = Right$("0" & Hex$(v), 2)
End Function

'------------------------------------------------------------------------------
' Keyboard shortcuts. The presets file stores readable text ("Ctrl+Alt+1") so a
' C# add-in can read it too; VBA converts it to Application.OnKey syntax.
'------------------------------------------------------------------------------

' "Ctrl+Alt+1" -> "^%1". Returns "" when the text is blank or invalid.
' A shortcut needs Ctrl or Alt (so plain typing is never hijacked), unless the
' key is a function key.
Public Function Core_ShortcutToOnKey(ByVal shortcut As String) As String
    Dim parts() As String, i As Long, token As String
    Dim mods As String, keyCode As String, hasCtrlOrAlt As Boolean, isFKey As Boolean

    shortcut = Trim$(shortcut)
    If Len(shortcut) = 0 Then Exit Function

    ' "Ctrl++" means Ctrl and the plus key.
    If Right$(shortcut, 2) = "++" Then shortcut = Left$(shortcut, Len(shortcut) - 2) & "+Plus"
    parts = Split(shortcut, "+")

    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If i < UBound(parts) Then
            Select Case LCase$(token)
                Case "ctrl", "control"
                    If InStr(mods, "^") > 0 Then Exit Function
                    mods = mods & "^": hasCtrlOrAlt = True
                Case "alt"
                    If InStr(mods, "%") > 0 Then Exit Function
                    mods = mods & "%": hasCtrlOrAlt = True
                Case "shift"
                    If InStr(mods, "+") > 0 Then Exit Function
                    mods = mods & "+"
                Case Else
                    Exit Function
            End Select
        Else
            keyCode = Core_KeyToOnKey(token, isFKey)
        End If
    Next i

    If Len(keyCode) = 0 Then Exit Function
    If Not hasCtrlOrAlt And Not isFKey Then Exit Function
    Core_ShortcutToOnKey = mods & keyCode
End Function

' Normalises how a shortcut is written: "ctrl + alt + a" -> "Ctrl+Alt+A".
' Returns "" when invalid.
Public Function Core_NormalizeShortcut(ByVal shortcut As String) As String
    Dim onKey As String, out As String, keyPart As String, i As Long, c As String
    onKey = Core_ShortcutToOnKey(shortcut)
    If Len(onKey) = 0 Then Exit Function

    ' Leading modifier characters, then the key.
    For i = 1 To Len(onKey)
        c = Mid$(onKey, i, 1)
        If c = "^" Or c = "%" Or (c = "+" And i < Len(onKey)) Then
            ' modifier
        Else
            keyPart = Mid$(onKey, i)
            Exit For
        End If
    Next i

    If InStr(Left$(onKey, Len(onKey) - Len(keyPart)), "^") > 0 Then out = out & "Ctrl+"
    If InStr(Left$(onKey, Len(onKey) - Len(keyPart)), "%") > 0 Then out = out & "Alt+"
    If InStr(Left$(onKey, Len(onKey) - Len(keyPart)), "+") > 0 Then out = out & "Shift+"

    If Left$(keyPart, 1) = "{" Then keyPart = Mid$(keyPart, 2, Len(keyPart) - 2)
    Select Case keyPart
        Case "+": keyPart = "Plus"
        Case "UP": keyPart = "Up"
        Case "DOWN": keyPart = "Down"
        Case "LEFT": keyPart = "Left"
        Case "RIGHT": keyPart = "Right"
        Case "HOME": keyPart = "Home"
        Case "END": keyPart = "End"
        Case "PGUP": keyPart = "PgUp"
        Case "PGDN": keyPart = "PgDn"
        Case "INSERT": keyPart = "Insert"
        Case "DEL": keyPart = "Delete"
        Case Else: keyPart = UCase$(keyPart)
    End Select
    Core_NormalizeShortcut = out & keyPart
End Function

Private Function Core_KeyToOnKey(ByVal token As String, ByRef isFKey As Boolean) As String
    Dim n As Long, u As String
    u = UCase$(token)
    isFKey = False

    If Len(token) = 1 Then
        Select Case token
            Case "A" To "Z", "a" To "z", "0" To "9"
                Core_KeyToOnKey = LCase$(token)
            Case "+", "^", "%", "~", "(", ")", "{", "}", "[", "]"
                Core_KeyToOnKey = "{" & token & "}"
            Case "-", "=", ";", "'", ",", ".", "/", "\", "`"
                Core_KeyToOnKey = token
        End Select
        Exit Function
    End If

    If Left$(u, 1) = "F" And IsNumeric(Mid$(u, 2)) Then
        n = CLng(Mid$(u, 2))
        If n >= 1 And n <= 15 Then
            Core_KeyToOnKey = "{F" & n & "}"
            isFKey = True
        End If
        Exit Function
    End If

    Select Case u
        Case "PLUS": Core_KeyToOnKey = "{+}"
        Case "UP": Core_KeyToOnKey = "{UP}"
        Case "DOWN": Core_KeyToOnKey = "{DOWN}"
        Case "LEFT": Core_KeyToOnKey = "{LEFT}"
        Case "RIGHT": Core_KeyToOnKey = "{RIGHT}"
        Case "HOME": Core_KeyToOnKey = "{HOME}"
        Case "END": Core_KeyToOnKey = "{END}"
        Case "PGUP", "PAGEUP": Core_KeyToOnKey = "{PGUP}"
        Case "PGDN", "PAGEDOWN": Core_KeyToOnKey = "{PGDN}"
        Case "INS", "INSERT": Core_KeyToOnKey = "{INSERT}"
        Case "DEL", "DELETE": Core_KeyToOnKey = "{DEL}"
    End Select
End Function

'------------------------------------------------------------------------------
' Text helpers
'------------------------------------------------------------------------------

' Escapes text for an XML attribute (used when building dynamicMenu content).
Public Function Core_XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, vbCr, "&#13;")      ' keep line breaks in super-tips
    s = Replace(s, vbLf, "&#10;")
    Core_XmlEscape = s
End Function

' Ribbon labels treat "&" as an accelerator prefix; double it to show it.
Public Function Core_RibbonLabel(ByVal s As String) As String
    Core_RibbonLabel = Replace(s, "&", "&&")
End Function

'------------------------------------------------------------------------------
' Collection helpers (VBA Collections cannot assign to an item in place)
'------------------------------------------------------------------------------

Public Sub Coll_ReplaceAt(ByVal coll As Collection, ByVal index As Long, ByVal value As Variant)
    If index > coll.Count Then
        coll.Add value
    Else
        coll.Add value, Before:=index
        coll.Remove index + 1
    End If
End Sub

Public Sub Coll_Swap(ByVal coll As Collection, ByVal i As Long, ByVal j As Long)
    Dim a As Variant, b As Variant
    If i = j Then Exit Sub
    If i < 1 Or j < 1 Or i > coll.Count Or j > coll.Count Then Exit Sub
    If IsObject(coll(i)) Then Set a = coll(i) Else a = coll(i)
    If IsObject(coll(j)) Then Set b = coll(j) Else b = coll(j)
    Coll_ReplaceAt coll, i, b
    Coll_ReplaceAt coll, j, a
End Sub

'------------------------------------------------------------------------------
' Versions
'------------------------------------------------------------------------------

' Compares "1.2.3" style versions (a leading "v" is ignored, missing parts count
' as 0, non-numeric parts as 0). Returns -1 when a < b, 0 when equal, 1 when a > b.
Public Function Core_CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim pa() As String, pb() As String, i As Long, n As Long, va As Long, vb As Long
    pa = Split(Core_StripV(a), ".")
    pb = Split(Core_StripV(b), ".")
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)
    For i = 0 To n
        va = 0: vb = 0
        If i <= UBound(pa) Then If IsNumeric(pa(i)) Then va = CLng(pa(i))
        If i <= UBound(pb) Then If IsNumeric(pb(i)) Then vb = CLng(pb(i))
        If va < vb Then Core_CompareVersions = -1: Exit Function
        If va > vb Then Core_CompareVersions = 1: Exit Function
    Next i
End Function

Private Function Core_StripV(ByVal s As String) As String
    s = Trim$(s)
    If LCase$(Left$(s, 1)) = "v" Then s = Mid$(s, 2)
    If Len(s) = 0 Then s = "0"
    Core_StripV = s
End Function

'------------------------------------------------------------------------------
' Colour by content: what kind of thing a formula is
'------------------------------------------------------------------------------

' Classifies a formula by where it gets its numbers from.
'   formula   the cell's formula text, starting with "=" (English; A1 or R1C1
'             notation both work). Anything else returns ckNone.
'   ownSheet  name of the sheet the cell is on.
'   names     LCase(name) -> the formula the name refers to. Sheet-scoped names
'             are keyed LCase(sheet) & "!" & LCase(name); tables are keyed by
'             LCase(table name) with a target such as "='Sheet'!A1". Can be
'             Nothing.
'   sheets    LCase(sheet name) -> True for every sheet in the workbook, so a
'             qualifier that is not a sheet (Book.xlsx!Name) reads as external.
'             The workbook's own name can be listed too, with any String as
'             its value: it then makes no link. Can be Nothing; every
'             unbracketed qualifier then counts as a sheet.
'
' Text in double quotes is never read as a reference. The strongest kind wins:
' external, then sheet link, then formula. A formula made only of numbers and
' operators (=1000*12) is an input.
Public Function Core_ClassifyFormula(ByVal formula As String, ByVal ownSheet As String, _
                                     ByVal names As Scripting.Dictionary, _
                                     ByVal sheets As Scripting.Dictionary) As DigitContentKind
    Core_ClassifyFormula = Cls_Scan(formula, LCase$(ownSheet), names, sheets, 0)
End Function

Private Function Cls_Scan(ByVal f As String, ByVal ownLower As String, ByVal names As Scripting.Dictionary, _
                          ByVal sheets As Scripting.Dictionary, ByVal depth As Long) As DigitContentKind
    Dim n As Long, i As Long, j As Long, k As Long, d As Long
    Dim ch As String, c As String, q As String, w As String, target As String
    Dim kind As DigitContentKind, bang As Boolean, was As Boolean, handled As Boolean

    If Left$(f, 1) <> "=" Then Exit Function
    n = Len(f)
    i = 2
    kind = ckInput

    Do While i <= n
        ch = Mid$(f, i, 1)
        was = bang                      ' the previous token was "Sheet!" or "[Book]Sheet!"
        bang = False

        If ch = """" Then
            ' Text literal: skip it ("" is an escaped quote).
            Cls_Raise kind, ckFormula
            i = i + 1
            Do While i <= n
                If Mid$(f, i, 1) = """" Then
                    If Mid$(f, i + 1, 1) = """" Then
                        i = i + 2
                    Else
                        Exit Do
                    End If
                Else
                    i = i + 1
                End If
            Loop
            i = i + 1

        ElseIf ch = "'" Then
            ' Quoted qualifier: 'My Sheet'!A1, 'C:\path\[Book.xlsx]Sheet1'!A1
            j = i + 1
            q = ""
            Do While j <= n
                If Mid$(f, j, 1) = "'" Then
                    If Mid$(f, j + 1, 1) = "'" Then
                        q = q & "'"
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    q = q & Mid$(f, j, 1)
                    j = j + 1
                End If
            Loop
            i = j + 1
            If Mid$(f, i, 1) = "!" Then
                Cls_Raise kind, Cls_Qualifier(q, ownLower, sheets)
                i = i + 1
                bang = True
            Else
                Cls_Raise kind, ckFormula
            End If

        ElseIf ch = "[" Then
            ' Either a workbook ([Book.xlsx]Sheet1!A1) or a structured / R1C1
            ' reference (Table1[Amount], [@Amount], RC[-1]). Only the first is
            ' followed by a sheet name and "!".
            d = 1
            j = i + 1
            Do While j <= n And d > 0
                c = Mid$(f, j, 1)
                If c = "'" Then
                    j = j + 2           ' ' escapes the next character in a column name
                Else
                    If c = "[" Then d = d + 1
                    If c = "]" Then d = d - 1
                    j = j + 1
                End If
            Loop
            i = j
            k = i
            Do While k <= n
                c = Mid$(f, k, 1)
                If Not (Cls_IsWord(c) Or c = ":") Then Exit Do
                k = k + 1
            Loop
            If k > i And Mid$(f, k, 1) = "!" Then
                Cls_Raise kind, ckExternal
                i = k + 1
                bang = True
            Else
                Cls_Raise kind, ckFormula
            End If

        ElseIf Cls_IsWordStart(ch) Then
            j = i
            Do While j <= n
                If Not Cls_IsWord(Mid$(f, j, 1)) Then Exit Do
                j = j + 1
            Loop
            w = Mid$(f, i, j - i)
            i = j
            c = Mid$(f, i, 1)
            handled = False

            If was Then
                ' The reference or name after "Sheet!": already accounted for.
                Cls_Raise kind, ckFormula
                handled = True
            ElseIf c = "!" Then
                Cls_Raise kind, Cls_Qualifier(w, ownLower, sheets)
                i = i + 1
                bang = True
                handled = True
            ElseIf c = ":" Then
                ' Jan:Dec!B5 is a 3D reference; A1:B2 is not.
                k = i + 1
                Do While k <= n
                    If Not Cls_IsWord(Mid$(f, k, 1)) Then Exit Do
                    k = k + 1
                Loop
                If k > i + 1 And Mid$(f, k, 1) = "!" Then
                    Cls_Raise kind, Cls_Qualifier(w & ":" & Mid$(f, i + 1, k - i - 1), ownLower, sheets)
                    i = k + 1
                    bang = True
                    handled = True
                End If
            End If

            If Not handled Then
                If c = "(" Then
                    Cls_Raise kind, ckFormula                   ' a function
                ElseIf StrComp(w, "TRUE", vbTextCompare) <> 0 And StrComp(w, "FALSE", vbTextCompare) <> 0 Then
                    ' A cell reference, a defined name, a table, or an unknown word.
                    Cls_Raise kind, ckFormula
                    If Not names Is Nothing And depth < 5 Then
                        target = ""
                        If names.Exists(ownLower & "!" & LCase$(w)) Then
                            target = names(ownLower & "!" & LCase$(w))
                        ElseIf names.Exists(LCase$(w)) Then
                            target = names(LCase$(w))
                        End If
                        If Len(target) > 0 Then Cls_Raise kind, Cls_Scan(target, ownLower, names, sheets, depth + 1)
                    End If
                End If
            End If

        ElseIf Cls_IsDigit(ch) Or (ch = "." And Cls_IsDigit(Mid$(f, i + 1, 1))) Then
            ' A number, with an optional exponent (1E+3).
            j = i
            Do While j <= n
                c = Mid$(f, j, 1)
                If Not (Cls_IsDigit(c) Or c = ".") Then Exit Do
                j = j + 1
            Loop
            c = Mid$(f, j, 1)
            If c = "E" Or c = "e" Then
                k = j + 1
                c = Mid$(f, k, 1)
                If c = "+" Or c = "-" Then k = k + 1
                If Cls_IsDigit(Mid$(f, k, 1)) Then
                    Do While Cls_IsDigit(Mid$(f, k, 1))
                        k = k + 1
                    Loop
                    j = k
                End If
            End If
            i = j

        ElseIf ch = "#" Then
            ' An error literal (#REF!, #N/A, #DIV/0!) or the spill operator (A1#).
            j = i + 1
            Do While j <= n
                c = Mid$(f, j, 1)
                If Not (Cls_IsLetter(c) Or c = "/" Or c = "0") Then Exit Do
                j = j + 1
            Loop
            If j > i + 1 Then
                Cls_Raise kind, ckFormula
                c = Mid$(f, j, 1)
                If c = "!" Or c = "?" Then j = j + 1
            End If
            i = j

        ElseIf ch = ":" Then
            Cls_Raise kind, ckFormula       ' a range such as 1:1
            i = i + 1

        Else
            i = i + 1
        End If
    Loop

    Cls_Scan = kind
End Function

' Kind of link a qualifier (the text before "!") makes: none when it is the
' cell's own sheet.
Private Function Cls_Qualifier(ByVal q As String, ByVal ownLower As String, _
                               ByVal sheets As Scripting.Dictionary) As DigitContentKind
    Dim p As Long, first As String

    ' A workbook in brackets or a path: sheet names cannot contain [ \ /.
    If InStr(q, "[") > 0 Or InStr(q, "\") > 0 Or InStr(q, "/") > 0 Then
        Cls_Qualifier = ckExternal
        Exit Function
    End If

    p = InStr(q, ":")                   ' Jan:Dec, a 3D reference
    If p > 0 Then first = Left$(q, p - 1) Else first = q
    first = LCase$(first)

    If Not sheets Is Nothing Then
        If Not sheets.Exists(first) Then
            Cls_Qualifier = ckExternal  ' Book.xlsx!Name, PERSONAL.XLSB!MyFn
            Exit Function
        End If
        If VarType(sheets(first)) = vbString Then
            Cls_Qualifier = ckFormula   ' this workbook's own name, not a sheet
            Exit Function
        End If
    End If

    If p = 0 And first = ownLower Then
        Cls_Qualifier = ckFormula
    Else
        Cls_Qualifier = ckSheetLink
    End If
End Function

' Raises kind to k when k is stronger. Anything that is not a plain number
' makes the formula at least ckFormula.
Private Sub Cls_Raise(ByRef kind As DigitContentKind, ByVal k As DigitContentKind)
    If k < ckFormula Then k = ckFormula
    If k > kind Then kind = k
End Sub

Private Function Cls_IsDigit(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Cls_IsDigit = (ch >= "0" And ch <= "9")
End Function

Private Function Cls_IsLetter(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Cls_IsLetter = (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")
End Function

' First character of a word: a letter, _ \ $ (as in $A$1), or any non-ASCII
' character (sheet and range names in other languages).
Private Function Cls_IsWordStart(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    If Cls_IsLetter(ch) Or ch = "_" Or ch = "\" Or ch = "$" Then
        Cls_IsWordStart = True
    Else
        Cls_IsWordStart = (AscW(ch) > 127 Or AscW(ch) < 0)
    End If
End Function

Private Function Cls_IsWord(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Cls_IsWord = Cls_IsWordStart(ch) Or Cls_IsDigit(ch) Or ch = "." Or ch = "?"
End Function
