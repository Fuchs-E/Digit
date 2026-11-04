Attribute VB_Name = "DigitStore"
'==============================================================================
' DigitStore - presets file: location, load, save, defaults and validation.
'
' Presets live outside the add-in in %APPDATA%\Digit\presets.json so they
' survive add-in updates and a future C# add-in can read the same file. The
' schema is described in docs/PRESETS.md.
'
' In memory the configuration is the parsed JSON tree (Scripting.Dictionary for
' objects, Collection for arrays), always passed through Store_Normalize so the
' rest of the code can rely on its shape:
'
'   cfg("version")                 Long
'   cfg("settings")("shortcutsEnabled")  Boolean
'   cfg("settings")("shortcuts")   Collection of 5 Strings ("Ctrl+Alt+1" or "")
'   cfg("settings")("undoDepth")   Long, 1..50
'   cfg("slots")                   Collection of 5 Strings (group name or "")
'   cfg("groups")                  Collection of Dictionary:
'       ("name")    String, unique (case-insensitive), non-empty
'       ("presets") Collection of 1..3 Dictionary:
'           ("numberFormat") String, non-empty
'           ("fontColor")    "#RRGGBB" or ""
'           ("fillColor")    "#RRGGBB" or ""
'
' Requires a reference to "Microsoft Scripting Runtime" (added by the build).
'==============================================================================
Option Explicit

Public Const DIGIT_FILE_VERSION As Long = 1

Private mCfg As Scripting.Dictionary
Private mPathOverride As String
Private mLastLoadMessage As String

'------------------------------------------------------------------------------
' Location
'------------------------------------------------------------------------------

Public Function Store_FilePath() As String
    If Len(mPathOverride) > 0 Then
        Store_FilePath = mPathOverride
    Else
        Store_FilePath = Store_Folder() & "\presets.json"
    End If
End Function

Public Function Store_Folder() As String
    If Len(mPathOverride) > 0 Then
        Store_Folder = Left$(mPathOverride, InStrRev(mPathOverride, "\") - 1)
    Else
        Store_Folder = Environ$("APPDATA") & "\Digit"
    End If
End Function

' Points the add-in at another presets file (used by the tests). Pass "" to reset.
Public Sub Store_SetPathOverride(ByVal path As String)
    mPathOverride = path
    Set mCfg = Nothing
End Sub

'------------------------------------------------------------------------------
' Load / save
'------------------------------------------------------------------------------

' The live configuration. Loads it on first use (after a VBA reset too).
Public Function Store_Config() As Scripting.Dictionary
    If mCfg Is Nothing Then Store_Load
    Set Store_Config = mCfg
End Function

' Text describing a problem found during the last load ("" when fine). The UI
' shows it once and clears it.
Public Function Store_TakeLoadMessage() As String
    Store_TakeLoadMessage = mLastLoadMessage
    mLastLoadMessage = ""
End Function

' Reads the presets file. Creates it with starter presets when missing. A file
' that cannot be parsed is renamed aside and replaced by the starter presets,
' so a typo never leaves the user with a broken ribbon.
Public Sub Store_Load()
    Dim path As String, text As String, errMsg As String, cfg As Scripting.Dictionary
    Dim aside As String

    path = Store_FilePath()
    mLastLoadMessage = ""

    If Len(Dir$(path)) = 0 Then
        Set mCfg = Store_DefaultConfig()
        If Not Store_WriteConfig(mCfg, errMsg) Then
            mLastLoadMessage = "Digit could not create its presets file:" & vbCrLf & path & vbCrLf & vbCrLf & errMsg
        End If
        Exit Sub
    End If

    If Not Store_ReadText(path, text, errMsg) Then
        Set mCfg = Store_DefaultConfig()
        mLastLoadMessage = "Digit could not read its presets file and is using starter presets for now:" & _
                           vbCrLf & path & vbCrLf & vbCrLf & errMsg
        Exit Sub
    End If

    Set cfg = Store_ParseConfig(text, errMsg)
    If cfg Is Nothing Then
        aside = Store_Folder() & "\presets.invalid-" & Format$(Now, "yyyymmdd-hhnnss") & ".json"
        On Error Resume Next
        Name path As aside
        On Error GoTo 0
        Set mCfg = Store_DefaultConfig()
        Store_WriteConfig mCfg, ""
        mLastLoadMessage = "Digit's presets file could not be read (" & errMsg & ")." & vbCrLf & vbCrLf & _
                           "It was renamed to:" & vbCrLf & aside & vbCrLf & vbCrLf & _
                           "Starter presets have been created in its place."
        Exit Sub
    End If

    Set mCfg = cfg
End Sub

' Parses and normalises presets JSON. Returns Nothing (and errMsg) when the text
' is not valid JSON or has no usable preset groups.
Public Function Store_ParseConfig(ByVal text As String, ByRef errMsg As String) As Scripting.Dictionary
    Dim parsed As Object, cfg As Scripting.Dictionary

    On Error GoTo ParseFail
    Set parsed = JsonConverter.ParseJson(text)
    On Error GoTo 0

    If TypeName(parsed) <> "Dictionary" Then
        errMsg = "the file does not contain a JSON object"
        Exit Function
    End If

    Set cfg = Store_Normalize(parsed)
    If cfg("groups").Count = 0 Then
        errMsg = "the file has no valid preset groups"
        Exit Function
    End If
    Set Store_ParseConfig = cfg
    Exit Function

ParseFail:
    errMsg = "invalid JSON: " & Err.Description
End Function

' Normalises and saves cfg as the live configuration.
Public Function Store_Save(ByVal cfg As Scripting.Dictionary, ByRef errMsg As String) As Boolean
    Dim clean As Scripting.Dictionary
    Set clean = Store_Normalize(cfg)
    If Store_WriteConfig(clean, errMsg) Then
        Set mCfg = clean
        Store_Save = True
    End If
End Function

' Writes cfg to the presets file. The previous file is kept as presets.json.bak
' and the new one is written to a temporary file first, so a failed write never
' leaves a half-written presets file behind.
Public Function Store_WriteConfig(ByVal cfg As Scripting.Dictionary, ByRef errMsg As String) As Boolean
    Dim path As String, tmp As String, bak As String
    path = Store_FilePath()
    tmp = path & ".tmp"
    bak = path & ".bak"

    If Not Store_EnsureFolder(Store_Folder(), errMsg) Then Exit Function
    If Not Store_WriteText(tmp, Store_ToJson(cfg), errMsg) Then Exit Function

    On Error GoTo Fail
    If Len(Dir$(path)) > 0 Then
        If Len(Dir$(bak)) > 0 Then Kill bak
        Name path As bak
    End If
    Name tmp As path
    Store_WriteConfig = True
    Exit Function

Fail:
    errMsg = Err.Description
End Function

Public Function Store_ToJson(ByVal cfg As Scripting.Dictionary) As String
    Store_ToJson = JsonConverter.ConvertToJson(cfg, 2)
End Function

' Deep copy, used by the settings window to edit without touching live presets.
Public Function Store_Clone(ByVal cfg As Scripting.Dictionary) As Scripting.Dictionary
    Set Store_Clone = Store_Normalize(JsonConverter.ParseJson(Store_ToJson(cfg)))
End Function

'------------------------------------------------------------------------------
' Queries and edits used by the ribbon, engine and settings window
'------------------------------------------------------------------------------

Public Function Store_Groups(ByVal cfg As Scripting.Dictionary) As Collection
    Set Store_Groups = cfg("groups")
End Function

' Office icon name for a group's ribbon button ("" = use the default).
Public Function Store_GroupIcon(ByVal g As Scripting.Dictionary) As String
    If g Is Nothing Then Exit Function
    If g.Exists("icon") Then Store_GroupIcon = CStr(g("icon"))
End Function

' Case-insensitive lookup. Returns Nothing when not found.
Public Function Store_FindGroup(ByVal cfg As Scripting.Dictionary, ByVal groupName As String) As Scripting.Dictionary
    Dim g As Scripting.Dictionary
    If Len(groupName) = 0 Then Exit Function
    For Each g In cfg("groups")
        If StrComp(g("name"), groupName, vbTextCompare) = 0 Then
            Set Store_FindGroup = g
            Exit Function
        End If
    Next g
End Function

Public Function Store_GroupIndex(ByVal cfg As Scripting.Dictionary, ByVal groupName As String) As Long
    Dim i As Long
    For i = 1 To cfg("groups").Count
        If StrComp(cfg("groups")(i)("name"), groupName, vbTextCompare) = 0 Then
            Store_GroupIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function Store_SlotGroupName(ByVal cfg As Scripting.Dictionary, ByVal slot As Long) As String
    If slot < 1 Or slot > DIGIT_SLOT_COUNT Then Exit Function
    Store_SlotGroupName = cfg("slots")(slot)
End Function

Public Sub Store_SetSlot(ByVal cfg As Scripting.Dictionary, ByVal slot As Long, ByVal groupName As String)
    If slot < 1 Or slot > DIGIT_SLOT_COUNT Then Exit Sub
    Coll_ReplaceAt cfg("slots"), slot, groupName
End Sub

' True when the group is on one of the quick buttons.
Public Function Store_IsInSlot(ByVal cfg As Scripting.Dictionary, ByVal groupName As String) As Boolean
    Dim i As Long
    For i = 1 To DIGIT_SLOT_COUNT
        If StrComp(cfg("slots")(i), groupName, vbTextCompare) = 0 Then
            Store_IsInSlot = True
            Exit Function
        End If
    Next i
End Function

Public Function Store_Shortcut(ByVal cfg As Scripting.Dictionary, ByVal slot As Long) As String
    If slot < 1 Or slot > DIGIT_SLOT_COUNT Then Exit Function
    Store_Shortcut = cfg("settings")("shortcuts")(slot)
End Function

Public Sub Store_SetShortcut(ByVal cfg As Scripting.Dictionary, ByVal slot As Long, ByVal shortcut As String)
    If slot < 1 Or slot > DIGIT_SLOT_COUNT Then Exit Sub
    Coll_ReplaceAt cfg("settings")("shortcuts"), slot, shortcut
End Sub

Public Function Store_ShortcutsEnabled(ByVal cfg As Scripting.Dictionary) As Boolean
    Store_ShortcutsEnabled = cfg("settings")("shortcutsEnabled")
End Function

Public Function Store_UndoDepth(ByVal cfg As Scripting.Dictionary) As Long
    Store_UndoDepth = cfg("settings")("undoDepth")
End Function

' Renames a group and keeps the quick-button assignments pointing at it.
Public Sub Store_RenameGroup(ByVal cfg As Scripting.Dictionary, ByVal oldName As String, ByVal newName As String)
    Dim g As Scripting.Dictionary, i As Long
    Set g = Store_FindGroup(cfg, oldName)
    If g Is Nothing Then Exit Sub
    g("name") = newName
    For i = 1 To DIGIT_SLOT_COUNT
        If StrComp(cfg("slots")(i), oldName, vbTextCompare) = 0 Then Store_SetSlot cfg, i, newName
    Next i
End Sub

' Deletes a group and clears any quick button that used it.
Public Sub Store_DeleteGroup(ByVal cfg As Scripting.Dictionary, ByVal groupName As String)
    Dim idx As Long, i As Long
    idx = Store_GroupIndex(cfg, groupName)
    If idx = 0 Then Exit Sub
    cfg("groups").Remove idx
    For i = 1 To DIGIT_SLOT_COUNT
        If StrComp(cfg("slots")(i), groupName, vbTextCompare) = 0 Then Store_SetSlot cfg, i, ""
    Next i
End Sub

Public Function Store_NewGroup(ByVal groupName As String) As Scripting.Dictionary
    Dim g As New Scripting.Dictionary
    g("name") = groupName
    g("icon") = ""
    Set g("presets") = New Collection
    Set Store_NewGroup = g
End Function

Public Function Store_NewPreset(ByVal numberFormat As String, Optional ByVal fontColor As String = "", _
                                Optional ByVal fillColor As String = "") As Scripting.Dictionary
    Dim p As New Scripting.Dictionary
    p("numberFormat") = numberFormat
    p("fontColor") = Core_NormalizeHex(fontColor)
    p("fillColor") = Core_NormalizeHex(fillColor)
    Set Store_NewPreset = p
End Function

' A group name not used yet: "New group", "New group 2", ...
Public Function Store_UniqueGroupName(ByVal cfg As Scripting.Dictionary, ByVal baseName As String) As String
    Dim n As Long, candidate As String
    candidate = baseName
    n = 1
    Do While Not Store_FindGroup(cfg, candidate) Is Nothing
        n = n + 1
        candidate = baseName & " " & n
    Loop
    Store_UniqueGroupName = candidate
End Function

'------------------------------------------------------------------------------
' Defaults
'------------------------------------------------------------------------------

' Starter presets written on first run. Nothing here is special: the user can
' rename, change or delete every group.
Public Function Store_DefaultConfig() As Scripting.Dictionary
    Dim cfg As New Scripting.Dictionary, settings As New Scripting.Dictionary
    Dim groups As New Collection, slots As New Collection, keys As New Collection
    Dim g As Scripting.Dictionary, i As Long

    Set g = Store_NewGroup("Numbers")
    g("presets").Add Store_NewPreset("#,##0;(#,##0);-")
    g("presets").Add Store_NewPreset("#,##0.0;(#,##0.0);-")
    g("presets").Add Store_NewPreset("#,##0.00;(#,##0.00);-")
    groups.Add g

    Set g = Store_NewGroup("Thousands")
    g("presets").Add Store_NewPreset("#,##0,;(#,##0,);-")
    g("presets").Add Store_NewPreset("#,##0.0,;(#,##0.0,);-")
    g("presets").Add Store_NewPreset("#,##0.0,,;(#,##0.0,,);-")
    groups.Add g

    Set g = Store_NewGroup("Percent")
    g("presets").Add Store_NewPreset("0%")
    g("presets").Add Store_NewPreset("0.0%")
    g("presets").Add Store_NewPreset("0.00%")
    groups.Add g

    Set g = Store_NewGroup("Currency")
    g("presets").Add Store_NewPreset("$#,##0;($#,##0);-")
    g("presets").Add Store_NewPreset("$#,##0.00;($#,##0.00);-")
    groups.Add g

    Set g = Store_NewGroup("Dates")
    g("presets").Add Store_NewPreset("dd-mmm-yy")
    g("presets").Add Store_NewPreset("mmm-yy")
    g("presets").Add Store_NewPreset("yyyy-mm-dd")
    groups.Add g

    Set g = Store_NewGroup("Multiples")
    g("presets").Add Store_NewPreset("0.0""x""")
    g("presets").Add Store_NewPreset("0.00""x""")
    groups.Add g

    Set g = Store_NewGroup("Inputs")
    g("presets").Add Store_NewPreset("#,##0;(#,##0);-", "#0000FF")
    g("presets").Add Store_NewPreset("#,##0.00;(#,##0.00);-", "#0000FF")
    groups.Add g

    slots.Add "Numbers"
    slots.Add "Thousands"
    slots.Add "Percent"
    slots.Add "Currency"
    slots.Add "Dates"

    For i = 1 To DIGIT_SLOT_COUNT
        keys.Add "Ctrl+Alt+" & i
    Next i

    settings("shortcutsEnabled") = True
    Set settings("shortcuts") = keys
    settings("undoDepth") = 10

    cfg("version") = DIGIT_FILE_VERSION
    Set cfg("settings") = settings
    Set cfg("slots") = slots
    Set cfg("groups") = groups
    Set Store_DefaultConfig = cfg
End Function

'------------------------------------------------------------------------------
' Validation / repair
'------------------------------------------------------------------------------

' Builds a clean configuration from anything that came out of the JSON parser.
' Unknown keys are dropped, missing ones get defaults, groups without a name or
' without presets are skipped, duplicate names get a suffix, presets beyond the
' third are dropped, invalid colours become blank, and slots that point at a
' missing group are cleared.
Public Function Store_Normalize(ByVal src As Object) As Scripting.Dictionary
    Dim cfg As New Scripting.Dictionary, settings As New Scripting.Dictionary
    Dim groups As New Collection, slots As New Collection, keys As New Collection
    Dim srcSettings As Object, srcGroups As Object, srcSlots As Object, srcKeys As Object
    Dim item As Variant, g As Scripting.Dictionary, p As Variant, presets As Collection
    Dim nm As String, fmt As String, i As Long, n As Long, slotName As String
    Dim found As Scripting.Dictionary, depth As Long

    ' --- groups
    Set srcGroups = DictObject(src, "groups")
    If TypeName(srcGroups) = "Collection" Then
        For Each item In srcGroups
            If TypeName(item) = "Dictionary" Then
                nm = Trim$(DictString(item, "name", ""))
                Set presets = New Collection
                If TypeName(DictObject(item, "presets")) = "Collection" Then
                    For Each p In item("presets")
                        If TypeName(p) = "Dictionary" Then
                            fmt = DictString(p, "numberFormat", "")
                            If Len(fmt) > 0 And presets.Count < DIGIT_MAX_PRESETS Then
                                presets.Add Store_NewPreset(fmt, DictString(p, "fontColor", ""), DictString(p, "fillColor", ""))
                            End If
                        End If
                    Next p
                End If
                If Len(nm) > 0 And presets.Count > 0 Then
                    If Not FindIn(groups, nm) Is Nothing Then
                        n = 2
                        Do While Not FindIn(groups, nm & " (" & n & ")") Is Nothing
                            n = n + 1
                        Loop
                        nm = nm & " (" & n & ")"
                    End If
                    Set g = New Scripting.Dictionary
                    g("name") = nm
                    g("icon") = Trim$(DictString(item, "icon", ""))
                    Set g("presets") = presets
                    groups.Add g
                End If
            End If
        Next item
    End If

    ' --- slots (group names; anything unknown becomes "")
    Set srcSlots = DictObject(src, "slots")
    For i = 1 To DIGIT_SLOT_COUNT
        slotName = ""
        If TypeName(srcSlots) = "Collection" Then
            If i <= srcSlots.Count Then
                If Not IsObject(srcSlots(i)) Then
                    If Not IsNull(srcSlots(i)) Then
                        Set found = FindIn(groups, CStr(srcSlots(i)))
                        If Not found Is Nothing Then slotName = found("name")
                    End If
                End If
            End If
        End If
        slots.Add slotName
    Next i

    ' --- settings
    Set srcSettings = DictObject(src, "settings")
    settings("shortcutsEnabled") = DictBool(srcSettings, "shortcutsEnabled", True)
    Set srcKeys = DictObject(srcSettings, "shortcuts")
    For i = 1 To DIGIT_SLOT_COUNT
        If TypeName(srcKeys) = "Collection" Then
            If i <= srcKeys.Count Then
                If IsObject(srcKeys(i)) Then
                    keys.Add ""
                ElseIf IsNull(srcKeys(i)) Then
                    keys.Add ""
                Else
                    keys.Add Core_NormalizeShortcut(CStr(srcKeys(i)))
                End If
            Else
                keys.Add ""
            End If
        Else
            keys.Add "Ctrl+Alt+" & i
        End If
    Next i
    Set settings("shortcuts") = keys
    depth = DictLong(srcSettings, "undoDepth", 10)
    If depth < 1 Then depth = 1
    If depth > 50 Then depth = 50
    settings("undoDepth") = depth

    cfg("version") = DIGIT_FILE_VERSION
    Set cfg("settings") = settings
    Set cfg("slots") = slots
    Set cfg("groups") = groups
    Set Store_Normalize = cfg
End Function

Private Function FindIn(ByVal groups As Collection, ByVal nm As String) As Scripting.Dictionary
    Dim g As Scripting.Dictionary
    For Each g In groups
        If StrComp(g("name"), nm, vbTextCompare) = 0 Then
            Set FindIn = g
            Exit Function
        End If
    Next g
End Function

Private Function DictObject(ByVal d As Object, ByVal key As String) As Object
    If d Is Nothing Then Exit Function
    If TypeName(d) <> "Dictionary" Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Then Set DictObject = d(key)
End Function

Private Function DictString(ByVal d As Object, ByVal key As String, ByVal def As String) As String
    DictString = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    DictString = CStr(d(key))
End Function

Private Function DictBool(ByVal d As Object, ByVal key As String, ByVal def As Boolean) As Boolean
    DictBool = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Then Exit Function
    If VarType(d(key)) = vbBoolean Then DictBool = d(key)
End Function

Private Function DictLong(ByVal d As Object, ByVal key As String, ByVal def As Long) As Long
    DictLong = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d(key)) Then Exit Function
    If IsNumeric(d(key)) Then DictLong = CLng(d(key))
End Function

'------------------------------------------------------------------------------
' File helpers (UTF-8 without BOM, so group names can use any language)
'------------------------------------------------------------------------------

Public Function Store_ReadText(ByVal path As String, ByRef text As String, ByRef errMsg As String) As Boolean
    Dim stm As Object
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2              ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    text = stm.ReadText(-1)   ' adReadAll
    stm.Close
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then text = Mid$(text, 2)
    End If
    Store_ReadText = True
    Exit Function
Fail:
    errMsg = Err.Description
End Function

Public Function Store_WriteText(ByVal path As String, ByVal text As String, ByRef errMsg As String) As Boolean
    Dim stm As Object, bin As Object
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2              ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    ' Skip the 3-byte BOM ADODB writes for utf-8.
    stm.Position = 3
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1              ' adTypeBinary
    bin.Open
    stm.CopyTo bin
    bin.SaveToFile path, 2    ' adSaveCreateOverWrite
    bin.Close
    stm.Close
    Store_WriteText = True
    Exit Function
Fail:
    errMsg = Err.Description
End Function

Public Function Store_EnsureFolder(ByVal folder As String, ByRef errMsg As String) As Boolean
    Dim parent As String
    If Len(Dir$(folder, vbDirectory)) > 0 Then
        Store_EnsureFolder = True
        Exit Function
    End If
    parent = Left$(folder, InStrRev(folder, "\") - 1)
    If Len(parent) > 2 Then
        If Not Store_EnsureFolder(parent, errMsg) Then Exit Function
    End If
    On Error GoTo Fail
    MkDir folder
    Store_EnsureFolder = True
    Exit Function
Fail:
    errMsg = "Cannot create folder " & folder & ": " & Err.Description
End Function
