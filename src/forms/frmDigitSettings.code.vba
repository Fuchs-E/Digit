'==============================================================================
' frmDigitSettings - preset editor.
'
' This is the form's code module. The form's controls are created by
' build/DigitFormBuilder.bas (control names below must match it), and build.ps1
' pastes this text into the form after the controls exist.
'
' The window edits a private copy of the presets; nothing is written until Save.
' The caller shows it modally and unloads it afterwards, so the form only ever
' hides itself.
'==============================================================================
Option Explicit

Private Const SAMPLE_DEFAULT As Double = 1234567.891

Private mCfg As Scripting.Dictionary     ' working copy
Private mLoading As Boolean              ' true while the code fills controls
Private mDirty As Boolean

'------------------------------------------------------------------------------
' Open / close
'------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    Me.Caption = "Digit " & DIGIT_VERSION & " - Presets"
    mLoading = True
    txtSample.Text = Format$(SAMPLE_DEFAULT, "0.000")
    FillFormatSuggestions
    FillIconSuggestions
    mLoading = False
    LoadAll Store_Clone(Store_Config())
    mDirty = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub

Private Sub btnSave_Click()
    Dim msg As String, i As Long
    If Not ValidateAll(msg) Then
        MsgBox msg, vbExclamation, DIGIT_TITLE
        Exit Sub
    End If
    For i = 1 To DIGIT_SLOT_COUNT
        Store_SetShortcut mCfg, i, Core_NormalizeShortcut(KeyBox(i).Text)
    Next i
    If Not Store_Save(mCfg, msg) Then
        MsgBox "Digit could not save the presets file:" & vbCrLf & Store_FilePath() & vbCrLf & vbCrLf & msg, _
               vbCritical, DIGIT_TITLE
        Exit Sub
    End If
    App_OnPresetsChanged
    mDirty = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    If mDirty Then
        If MsgBox("Discard your changes to the presets?", vbQuestion + vbYesNo + vbDefaultButton2, DIGIT_TITLE) <> vbYes Then
            Exit Sub
        End If
    End If
    Me.Hide
End Sub

' Fills every control from cfg (used on open and after an import).
Private Sub LoadAll(ByVal cfg As Scripting.Dictionary)
    Dim i As Long
    Set mCfg = cfg
    mLoading = True
    chkShortcuts.Value = Store_ShortcutsEnabled(mCfg)
    For i = 1 To DIGIT_SLOT_COUNT
        KeyBox(i).Text = Store_Shortcut(mCfg, i)
        KeyBox(i).ForeColor = vbWindowText
    Next i
    mLoading = False
    RefreshGroups 1
    RefreshSlots
    lblStatus.Caption = "File: " & Store_FilePath()
    lblStatus.ControlTipText = Store_FilePath()
End Sub

'------------------------------------------------------------------------------
' Groups
'------------------------------------------------------------------------------

Private Function CurGroup() As Scripting.Dictionary
    If lstGroups.ListIndex < 0 Then Exit Function
    If lstGroups.ListIndex + 1 > mCfg("groups").Count Then Exit Function
    Set CurGroup = mCfg("groups")(lstGroups.ListIndex + 1)
End Function

Private Sub RefreshGroups(ByVal selectIndex As Long)
    Dim g As Scripting.Dictionary
    mLoading = True
    lstGroups.Clear
    For Each g In mCfg("groups")
        lstGroups.AddItem g("name")
    Next g
    If lstGroups.ListCount > 0 Then
        If selectIndex < 1 Then selectIndex = 1
        If selectIndex > lstGroups.ListCount Then selectIndex = lstGroups.ListCount
        lstGroups.ListIndex = selectIndex - 1
    End If
    mLoading = False
    RefreshPresets 1
End Sub

Private Sub lstGroups_Click()
    If mLoading Then Exit Sub
    RefreshPresets 1
End Sub

Private Sub btnGroupNew_Click()
    Dim nm As String, g As Scripting.Dictionary
    nm = AskGroupName("Name for the new preset group:", Store_UniqueGroupName(mCfg, "New group"), "")
    If Len(nm) = 0 Then Exit Sub
    Set g = Store_NewGroup(nm)
    g("presets").Add Store_NewPreset("#,##0")
    mCfg("groups").Add g
    mDirty = True
    RefreshGroups mCfg("groups").Count
    RefreshSlots
    On Error Resume Next
    cboFormat.SetFocus
End Sub

Private Sub btnGroupRename_Click()
    Dim g As Scripting.Dictionary, oldName As String, nm As String
    Set g = CurGroup()
    If g Is Nothing Then Exit Sub
    oldName = g("name")
    nm = AskGroupName("New name for """ & oldName & """:", oldName, oldName)
    If Len(nm) = 0 Or nm = oldName Then Exit Sub
    Store_RenameGroup mCfg, oldName, nm
    mDirty = True
    RefreshGroups lstGroups.ListIndex + 1
    RefreshSlots
End Sub

Private Sub btnGroupDelete_Click()
    Dim g As Scripting.Dictionary, idx As Long
    Set g = CurGroup()
    If g Is Nothing Then Exit Sub
    If mCfg("groups").Count = 1 Then
        MsgBox "Digit needs at least one preset group.", vbInformation, DIGIT_TITLE
        Exit Sub
    End If
    If MsgBox("Delete the preset group """ & g("name") & """?", vbQuestion + vbYesNo + vbDefaultButton2, DIGIT_TITLE) <> vbYes Then
        Exit Sub
    End If
    idx = lstGroups.ListIndex + 1
    Store_DeleteGroup mCfg, g("name")
    mDirty = True
    RefreshGroups idx
    RefreshSlots
End Sub

Private Sub btnGroupUp_Click()
    MoveGroup -1
End Sub

Private Sub btnGroupDown_Click()
    MoveGroup 1
End Sub

Private Sub MoveGroup(ByVal delta As Long)
    Dim i As Long
    i = lstGroups.ListIndex + 1
    If i < 1 Or i + delta < 1 Or i + delta > mCfg("groups").Count Then Exit Sub
    Coll_Swap mCfg("groups"), i, i + delta
    mDirty = True
    RefreshGroups i + delta
    RefreshSlots
End Sub

' Asks for a group name until it is unique (or the user cancels -> "").
Private Function AskGroupName(ByVal prompt As String, ByVal defaultName As String, ByVal currentName As String) As String
    Dim s As String, existing As Scripting.Dictionary
    Do
        s = Trim$(InputBox(prompt, DIGIT_TITLE, defaultName))
        If Len(s) = 0 Then Exit Function
        Set existing = Store_FindGroup(mCfg, s)
        If existing Is Nothing Or StrComp(s, currentName, vbTextCompare) = 0 Then
            AskGroupName = s
            Exit Function
        End If
        MsgBox "There is already a group called """ & existing("name") & """.", vbExclamation, DIGIT_TITLE
        defaultName = s
    Loop
End Function

'------------------------------------------------------------------------------
' Presets of the selected group
'------------------------------------------------------------------------------

Private Function CurPresetIndex() As Long
    CurPresetIndex = lstPresets.ListIndex + 1
End Function

Private Function CurPreset() As Scripting.Dictionary
    Dim g As Scripting.Dictionary, i As Long
    Set g = CurGroup()
    If g Is Nothing Then Exit Function
    i = CurPresetIndex()
    If i < 1 Or i > g("presets").Count Then Exit Function
    Set CurPreset = g("presets")(i)
End Function

Private Sub RefreshPresets(ByVal selectIndex As Long)
    Dim g As Scripting.Dictionary, p As Scripting.Dictionary, i As Long, v As Double
    mLoading = True
    lstPresets.Clear
    Set g = CurGroup()
    If Not g Is Nothing Then
        v = SampleValue()
        For i = 1 To g("presets").Count
            Set p = g("presets")(i)
            lstPresets.AddItem CStr(i)
            lstPresets.List(i - 1, 1) = p("numberFormat")
            lstPresets.List(i - 1, 2) = Fmt_Preview(p("numberFormat"), v)
        Next i
        If lstPresets.ListCount > 0 Then
            If selectIndex < 1 Then selectIndex = 1
            If selectIndex > lstPresets.ListCount Then selectIndex = lstPresets.ListCount
            lstPresets.ListIndex = selectIndex - 1
        End If
    End If
    mLoading = False
    LoadIcon
    LoadPreset
End Sub

Private Sub lstPresets_Click()
    If mLoading Then Exit Sub
    LoadPreset
End Sub

Private Sub LoadPreset()
    Dim p As Scripting.Dictionary
    Set p = CurPreset()
    mLoading = True
    If p Is Nothing Then
        cboFormat.Text = ""
    Else
        cboFormat.Text = p("numberFormat")
    End If
    mLoading = False
    UpdateEditor
End Sub

' Swatches, preview, validation and button states for the selected preset.
Private Sub UpdateEditor()
    Dim g As Scripting.Dictionary, p As Scripting.Dictionary, idx As Long, n As Long
    Dim fmt As String, ok As Boolean, v As Double, fc As Long, bc As Long

    Set g = CurGroup()
    Set p = CurPreset()
    idx = CurPresetIndex()
    If Not g Is Nothing Then n = g("presets").Count

    btnGroupRename.Enabled = Not g Is Nothing
    btnGroupDelete.Enabled = Not g Is Nothing
    btnGroupUp.Enabled = lstGroups.ListIndex > 0
    btnGroupDown.Enabled = lstGroups.ListIndex >= 0 And lstGroups.ListIndex < lstGroups.ListCount - 1

    btnPresetAdd.Enabled = Not g Is Nothing And n < DIGIT_MAX_PRESETS
    btnPresetDelete.Enabled = Not p Is Nothing And n > 1
    btnPresetUp.Enabled = Not p Is Nothing And idx > 1
    btnPresetDown.Enabled = Not p Is Nothing And idx < n

    cboFormat.Enabled = Not p Is Nothing
    btnFontPick.Enabled = Not p Is Nothing
    btnFontNone.Enabled = Not p Is Nothing
    btnFillPick.Enabled = Not p Is Nothing
    btnFillNone.Enabled = Not p Is Nothing

    If p Is Nothing Then
        ShowSwatch lblFontSwatch, ""
        ShowSwatch lblFillSwatch, ""
        lblFormatStatus.Caption = ""
        lblPreview.Caption = ""
        Exit Sub
    End If

    ShowSwatch lblFontSwatch, p("fontColor")
    ShowSwatch lblFillSwatch, p("fillColor")

    fmt = p("numberFormat")
    ok = Fmt_IsValidNumberFormat(fmt)
    If ok Then
        lblFormatStatus.Caption = "Valid"
        lblFormatStatus.ForeColor = RGB(0, 128, 0)
        v = SampleValue()
        lblPreview.Caption = Fmt_Preview(fmt, v) & "      " & Fmt_Preview(fmt, -v) & "      " & Fmt_Preview(fmt, 0)
    ElseIf Len(fmt) = 0 Then
        lblFormatStatus.Caption = "Required"
        lblFormatStatus.ForeColor = RGB(192, 0, 0)
        lblPreview.Caption = ""
    Else
        lblFormatStatus.Caption = "Invalid format"
        lblFormatStatus.ForeColor = RGB(192, 0, 0)
        lblPreview.Caption = "(Excel does not accept this number format)"
    End If

    fc = Core_HexToColor(p("fontColor"))
    bc = Core_HexToColor(p("fillColor"))
    If fc >= 0 Then lblPreview.ForeColor = fc Else lblPreview.ForeColor = vbBlack
    If bc >= 0 Then lblPreview.BackColor = bc Else lblPreview.BackColor = vbWhite
End Sub

Private Sub ShowSwatch(ByVal lbl As MSForms.Label, ByVal hexColor As String)
    Dim c As Long
    c = Core_HexToColor(hexColor)
    If c >= 0 Then
        lbl.BackColor = c
        lbl.Caption = ""
        lbl.ControlTipText = hexColor
    Else
        lbl.BackColor = vbButtonFace
        lbl.Caption = "none"
        lbl.ControlTipText = "Leave as is"
    End If
End Sub

Private Sub cboFormat_Change()
    Dim p As Scripting.Dictionary, i As Long
    If mLoading Then Exit Sub
    Set p = CurPreset()
    If p Is Nothing Then Exit Sub
    p("numberFormat") = cboFormat.Text
    mDirty = True
    i = lstPresets.ListIndex
    If i >= 0 Then
        lstPresets.List(i, 1) = cboFormat.Text
        lstPresets.List(i, 2) = Fmt_Preview(cboFormat.Text, SampleValue())
    End If
    UpdateEditor
End Sub

Private Sub btnPresetAdd_Click()
    Dim g As Scripting.Dictionary, p As Scripting.Dictionary
    Set g = CurGroup()
    If g Is Nothing Then Exit Sub
    If g("presets").Count >= DIGIT_MAX_PRESETS Then Exit Sub
    Set p = CurPreset()
    If p Is Nothing Then
        g("presets").Add Store_NewPreset("General")
    Else
        g("presets").Add Store_NewPreset(p("numberFormat"), p("fontColor"), p("fillColor"))
    End If
    mDirty = True
    RefreshPresets g("presets").Count
    On Error Resume Next
    cboFormat.SetFocus
End Sub

Private Sub btnPresetDelete_Click()
    Dim g As Scripting.Dictionary, i As Long
    Set g = CurGroup()
    i = CurPresetIndex()
    If g Is Nothing Or i < 1 Then Exit Sub
    If g("presets").Count <= 1 Then Exit Sub
    g("presets").Remove i
    mDirty = True
    RefreshPresets i
End Sub

Private Sub btnPresetUp_Click()
    MovePreset -1
End Sub

Private Sub btnPresetDown_Click()
    MovePreset 1
End Sub

Private Sub MovePreset(ByVal delta As Long)
    Dim g As Scripting.Dictionary, i As Long
    Set g = CurGroup()
    i = CurPresetIndex()
    If g Is Nothing Then Exit Sub
    If i < 1 Or i + delta < 1 Or i + delta > g("presets").Count Then Exit Sub
    Coll_Swap g("presets"), i, i + delta
    mDirty = True
    RefreshPresets i + delta
End Sub

'------------------------------------------------------------------------------
' Colours
'------------------------------------------------------------------------------

Private Sub btnFontPick_Click()
    PickColour "fontColor"
End Sub

Private Sub lblFontSwatch_Click()
    If btnFontPick.Enabled Then PickColour "fontColor"
End Sub

Private Sub btnFontNone_Click()
    ClearColour "fontColor"
End Sub

Private Sub btnFillPick_Click()
    PickColour "fillColor"
End Sub

Private Sub lblFillSwatch_Click()
    If btnFillPick.Enabled Then PickColour "fillColor"
End Sub

Private Sub btnFillNone_Click()
    ClearColour "fillColor"
End Sub

Private Sub PickColour(ByVal key As String)
    Dim p As Scripting.Dictionary, c As Long
    Set p = CurPreset()
    If p Is Nothing Then Exit Sub
    c = UI_PickColor(Core_HexToColor(p(key)), Me.Caption)
    If c < 0 Then Exit Sub
    p(key) = Core_ColorToHex(c)
    mDirty = True
    UpdateEditor
End Sub

Private Sub ClearColour(ByVal key As String)
    Dim p As Scripting.Dictionary
    Set p = CurPreset()
    If p Is Nothing Then Exit Sub
    p(key) = ""
    mDirty = True
    UpdateEditor
End Sub

'------------------------------------------------------------------------------
' Preview sample
'------------------------------------------------------------------------------

Private Function SampleValue() As Double
    If IsNumeric(txtSample.Text) Then
        SampleValue = CDbl(txtSample.Text)
    Else
        SampleValue = SAMPLE_DEFAULT
    End If
End Function

Private Sub txtSample_Change()
    Dim g As Scripting.Dictionary, i As Long, v As Double
    If mLoading Then Exit Sub
    Set g = CurGroup()
    If g Is Nothing Then Exit Sub
    v = SampleValue()
    For i = 1 To g("presets").Count
        If i <= lstPresets.ListCount Then lstPresets.List(i - 1, 2) = Fmt_Preview(g("presets")(i)("numberFormat"), v)
    Next i
    UpdateEditor
End Sub

Private Sub FillFormatSuggestions()
    Dim f As Variant
    cboFormat.Clear
    For Each f In Array("General", "0", "0.0", "0.00", "#,##0", "#,##0.00", _
                        "#,##0;(#,##0);-", "#,##0.0;(#,##0.0);-", "#,##0.00;(#,##0.00);-", _
                        "#,##0,;(#,##0,);-", "#,##0.0,;(#,##0.0,);-", "#,##0.0,,;(#,##0.0,,);-", _
                        "0%", "0.0%", "0.00%", "0.0%;(0.0%);-", _
                        "$#,##0;($#,##0);-", "$#,##0.00;($#,##0.00);-", _
                        "0.0""x""", "0.00""x""", "0.00E+00", _
                        "dd-mmm-yy", "mmm-yy", "yyyy-mm-dd", "d mmm yyyy", "hh:mm", "@")
        cboFormat.AddItem f
    Next f
End Sub

'------------------------------------------------------------------------------
' Button picture (Office icon name) of the selected group
'------------------------------------------------------------------------------

Private Const ICON_CELLS As Long = 66

' Fills the picker grid with the suggested icons this Office version has.
Private Sub FillIconSuggestions()
    Dim nm As Variant, k As Long, cell As MSForms.Image
    For Each nm In Split(DIGIT_ICON_SUGGESTIONS, ",")
        If k >= ICON_CELLS Then Exit For
        If UI_IsOfficeIcon(CStr(nm)) Then
            k = k + 1
            Set cell = Me.Controls("imgPick" & k)
            Set cell.Picture = UI_OfficeIconPicture(CStr(nm))
            cell.Tag = CStr(nm)
            cell.ControlTipText = CStr(nm)
        End If
    Next nm
    For k = k + 1 To ICON_CELLS
        Me.Controls("imgPick" & k).Visible = False
    Next k
End Sub

Private Sub LoadIcon()
    Dim g As Scripting.Dictionary
    Set g = CurGroup()
    mLoading = True
    txtIcon.Text = Store_GroupIcon(g)
    mLoading = False
    btnIconPick.Enabled = Not g Is Nothing
    btnIconNone.Enabled = Not g Is Nothing
    ShowIcon
End Sub

Private Sub SetIcon(ByVal nm As String)
    Dim g As Scripting.Dictionary
    Set g = CurGroup()
    If g Is Nothing Then Exit Sub
    g("icon") = Trim$(nm)
    mDirty = True
    mLoading = True
    txtIcon.Text = Trim$(nm)
    mLoading = False
    ShowIcon
End Sub

Private Sub btnIconPick_Click()
    If CurGroup() Is Nothing Then Exit Sub
    MarkPicked
    fraIcons.Visible = True
    fraIcons.ZOrder 0
End Sub

Private Sub imgIcon_Click()
    btnIconPick_Click
End Sub

Private Sub btnIconNone_Click()
    SetIcon ""
End Sub

Private Sub btnIconClose_Click()
    fraIcons.Visible = False
End Sub

Private Sub PickIcon(ByVal k As Long)
    SetIcon Me.Controls("imgPick" & k).Tag
    fraIcons.Visible = False
End Sub

Private Sub txtIcon_Change()
    If mLoading Then Exit Sub
    SetIcon txtIcon.Text
    MarkPicked
End Sub

' Frames the current picture in the grid.
Private Sub MarkPicked()
    Dim k As Long, cell As MSForms.Image
    For k = 1 To ICON_CELLS
        Set cell = Me.Controls("imgPick" & k)
        If cell.Visible Then
            If Len(cell.Tag) > 0 And StrComp(cell.Tag, Trim$(txtIcon.Text), vbTextCompare) = 0 Then
                cell.BorderStyle = 1
                cell.BackColor = &HF7E6CC               ' light blue
            Else
                cell.BorderStyle = 0
                cell.BackColor = &HFFFFFF
            End If
        End If
    Next k
End Sub

' Preview of the picture; a name Office does not know shows in red.
Private Sub ShowIcon()
    Dim nm As String
    nm = Trim$(txtIcon.Text)
    If Len(nm) = 0 Then
        Set imgIcon.Picture = Nothing
        txtIcon.ForeColor = vbWindowText
        imgIcon.ControlTipText = "No picture: the ribbon shows the button number. Click to choose one."
    ElseIf UI_IsOfficeIcon(nm) Then
        Set imgIcon.Picture = UI_OfficeIconPicture(nm)
        txtIcon.ForeColor = vbWindowText
        imgIcon.ControlTipText = nm & " (click to change)"
    Else
        Set imgIcon.Picture = Nothing
        txtIcon.ForeColor = vbRed
        imgIcon.ControlTipText = "Office has no icon called """ & nm & """."
    End If
End Sub

Private Sub imgPick1_Click(): PickIcon 1: End Sub
Private Sub imgPick2_Click(): PickIcon 2: End Sub
Private Sub imgPick3_Click(): PickIcon 3: End Sub
Private Sub imgPick4_Click(): PickIcon 4: End Sub
Private Sub imgPick5_Click(): PickIcon 5: End Sub
Private Sub imgPick6_Click(): PickIcon 6: End Sub
Private Sub imgPick7_Click(): PickIcon 7: End Sub
Private Sub imgPick8_Click(): PickIcon 8: End Sub
Private Sub imgPick9_Click(): PickIcon 9: End Sub
Private Sub imgPick10_Click(): PickIcon 10: End Sub
Private Sub imgPick11_Click(): PickIcon 11: End Sub
Private Sub imgPick12_Click(): PickIcon 12: End Sub
Private Sub imgPick13_Click(): PickIcon 13: End Sub
Private Sub imgPick14_Click(): PickIcon 14: End Sub
Private Sub imgPick15_Click(): PickIcon 15: End Sub
Private Sub imgPick16_Click(): PickIcon 16: End Sub
Private Sub imgPick17_Click(): PickIcon 17: End Sub
Private Sub imgPick18_Click(): PickIcon 18: End Sub
Private Sub imgPick19_Click(): PickIcon 19: End Sub
Private Sub imgPick20_Click(): PickIcon 20: End Sub
Private Sub imgPick21_Click(): PickIcon 21: End Sub
Private Sub imgPick22_Click(): PickIcon 22: End Sub
Private Sub imgPick23_Click(): PickIcon 23: End Sub
Private Sub imgPick24_Click(): PickIcon 24: End Sub
Private Sub imgPick25_Click(): PickIcon 25: End Sub
Private Sub imgPick26_Click(): PickIcon 26: End Sub
Private Sub imgPick27_Click(): PickIcon 27: End Sub
Private Sub imgPick28_Click(): PickIcon 28: End Sub
Private Sub imgPick29_Click(): PickIcon 29: End Sub
Private Sub imgPick30_Click(): PickIcon 30: End Sub
Private Sub imgPick31_Click(): PickIcon 31: End Sub
Private Sub imgPick32_Click(): PickIcon 32: End Sub
Private Sub imgPick33_Click(): PickIcon 33: End Sub
Private Sub imgPick34_Click(): PickIcon 34: End Sub
Private Sub imgPick35_Click(): PickIcon 35: End Sub
Private Sub imgPick36_Click(): PickIcon 36: End Sub
Private Sub imgPick37_Click(): PickIcon 37: End Sub
Private Sub imgPick38_Click(): PickIcon 38: End Sub
Private Sub imgPick39_Click(): PickIcon 39: End Sub
Private Sub imgPick40_Click(): PickIcon 40: End Sub
Private Sub imgPick41_Click(): PickIcon 41: End Sub
Private Sub imgPick42_Click(): PickIcon 42: End Sub
Private Sub imgPick43_Click(): PickIcon 43: End Sub
Private Sub imgPick44_Click(): PickIcon 44: End Sub
Private Sub imgPick45_Click(): PickIcon 45: End Sub
Private Sub imgPick46_Click(): PickIcon 46: End Sub
Private Sub imgPick47_Click(): PickIcon 47: End Sub
Private Sub imgPick48_Click(): PickIcon 48: End Sub
Private Sub imgPick49_Click(): PickIcon 49: End Sub
Private Sub imgPick50_Click(): PickIcon 50: End Sub
Private Sub imgPick51_Click(): PickIcon 51: End Sub
Private Sub imgPick52_Click(): PickIcon 52: End Sub
Private Sub imgPick53_Click(): PickIcon 53: End Sub
Private Sub imgPick54_Click(): PickIcon 54: End Sub
Private Sub imgPick55_Click(): PickIcon 55: End Sub
Private Sub imgPick56_Click(): PickIcon 56: End Sub
Private Sub imgPick57_Click(): PickIcon 57: End Sub
Private Sub imgPick58_Click(): PickIcon 58: End Sub
Private Sub imgPick59_Click(): PickIcon 59: End Sub
Private Sub imgPick60_Click(): PickIcon 60: End Sub
Private Sub imgPick61_Click(): PickIcon 61: End Sub
Private Sub imgPick62_Click(): PickIcon 62: End Sub
Private Sub imgPick63_Click(): PickIcon 63: End Sub
Private Sub imgPick64_Click(): PickIcon 64: End Sub
Private Sub imgPick65_Click(): PickIcon 65: End Sub
Private Sub imgPick66_Click(): PickIcon 66: End Sub

'------------------------------------------------------------------------------
' Quick-button slots and shortcuts
'------------------------------------------------------------------------------

Private Function SlotBox(ByVal i As Long) As MSForms.ComboBox
    Set SlotBox = Me.Controls("cboSlot" & i)
End Function

Private Function KeyBox(ByVal i As Long) As MSForms.TextBox
    Set KeyBox = Me.Controls("txtKey" & i)
End Function

Private Sub RefreshSlots()
    Dim i As Long, cb As MSForms.ComboBox, g As Scripting.Dictionary, cur As String
    mLoading = True
    For i = 1 To DIGIT_SLOT_COUNT
        Set cb = SlotBox(i)
        cb.Clear
        cb.AddItem "(none)"
        For Each g In mCfg("groups")
            cb.AddItem g("name")
        Next g
        cur = Store_SlotGroupName(mCfg, i)
        If Len(cur) = 0 Then
            cb.ListIndex = 0
        Else
            cb.ListIndex = Store_GroupIndex(mCfg, cur)
        End If
    Next i
    mLoading = False
    UpdateMoreHint
End Sub

Private Sub SlotChanged(ByVal i As Long)
    Dim cb As MSForms.ComboBox
    If mLoading Then Exit Sub
    Set cb = SlotBox(i)
    If cb.ListIndex <= 0 Then
        Store_SetSlot mCfg, i, ""
    Else
        Store_SetSlot mCfg, i, cb.List(cb.ListIndex)
    End If
    mDirty = True
    UpdateMoreHint
End Sub

Private Sub UpdateMoreHint()
    Dim g As Scripting.Dictionary, s As String
    For Each g In mCfg("groups")
        If Not Store_IsInSlot(mCfg, g("name")) Then
            If Len(s) > 0 Then s = s & ", "
            s = s & g("name")
        End If
    Next g
    If Len(s) = 0 Then s = "(none)"
    lblHint.Caption = "Under ""More presets"": " & s
End Sub

Private Sub cboSlot1_Change(): SlotChanged 1: End Sub
Private Sub cboSlot2_Change(): SlotChanged 2: End Sub
Private Sub cboSlot3_Change(): SlotChanged 3: End Sub
Private Sub cboSlot4_Change(): SlotChanged 4: End Sub
Private Sub cboSlot5_Change(): SlotChanged 5: End Sub

Private Sub KeyChanged(ByVal i As Long)
    Dim t As String
    If mLoading Then Exit Sub
    t = Trim$(KeyBox(i).Text)
    Store_SetShortcut mCfg, i, t
    If Len(t) > 0 And Len(Core_ShortcutToOnKey(t)) = 0 Then
        KeyBox(i).ForeColor = RGB(192, 0, 0)
    Else
        KeyBox(i).ForeColor = vbWindowText
    End If
    mDirty = True
End Sub

Private Sub txtKey1_Change(): KeyChanged 1: End Sub
Private Sub txtKey2_Change(): KeyChanged 2: End Sub
Private Sub txtKey3_Change(): KeyChanged 3: End Sub
Private Sub txtKey4_Change(): KeyChanged 4: End Sub
Private Sub txtKey5_Change(): KeyChanged 5: End Sub

Private Sub chkShortcuts_Click()
    If mLoading Then Exit Sub
    mCfg("settings")("shortcutsEnabled") = CBool(chkShortcuts.Value)
    mDirty = True
End Sub

'------------------------------------------------------------------------------
' Validation
'------------------------------------------------------------------------------

' Checks every number format and shortcut. On failure selects the culprit and
' returns a message.
Private Function ValidateAll(ByRef msg As String) As Boolean
    Dim gi As Long, pi As Long, g As Scripting.Dictionary, fmt As String
    Dim i As Long, j As Long, ki As String, kj As String

    If mCfg("groups").Count = 0 Then
        msg = "Add at least one preset group."
        Exit Function
    End If

    For gi = 1 To mCfg("groups").Count
        Set g = mCfg("groups")(gi)
        If Len(Store_GroupIcon(g)) > 0 And Not UI_IsOfficeIcon(Store_GroupIcon(g)) Then
            RefreshGroups gi
            msg = "The button picture for """ & g("name") & """ (""" & Store_GroupIcon(g) & """) is not an Office icon name." & _
                  vbCrLf & vbCrLf & "Pick one from the list, or leave it blank to show the button number."
            Exit Function
        End If
        For pi = 1 To g("presets").Count
            fmt = g("presets")(pi)("numberFormat")
            If Not Fmt_IsValidNumberFormat(fmt) Then
                RefreshGroups gi
                RefreshPresets pi
                If Len(fmt) = 0 Then
                    msg = "Preset " & pi & " of """ & g("name") & """ needs a number format."
                Else
                    msg = "Preset " & pi & " of """ & g("name") & """ has a number format Excel does not accept:" & _
                          vbCrLf & vbCrLf & fmt
                End If
                Exit Function
            End If
        Next pi
    Next gi

    For i = 1 To DIGIT_SLOT_COUNT
        ki = Trim$(KeyBox(i).Text)
        If Len(ki) > 0 And Len(Core_ShortcutToOnKey(ki)) = 0 Then
            msg = "The shortcut for button " & i & " (""" & ki & """) is not valid." & vbCrLf & vbCrLf & _
                  "Use Ctrl and/or Alt (optionally Shift) plus a key, for example Ctrl+Alt+1 or Ctrl+Shift+F5."
            Exit Function
        End If
    Next i

    For i = 1 To DIGIT_SLOT_COUNT - 1
        ki = Core_ShortcutToOnKey(KeyBox(i).Text)
        For j = i + 1 To DIGIT_SLOT_COUNT
            kj = Core_ShortcutToOnKey(KeyBox(j).Text)
            If Len(ki) > 0 And ki = kj Then
                msg = "Buttons " & i & " and " & j & " use the same shortcut."
                Exit Function
            End If
        Next j
    Next i

    ValidateAll = True
End Function

'------------------------------------------------------------------------------
' Import / export
'------------------------------------------------------------------------------

Private Sub btnImport_Click()
    Dim f As Variant, text As String, errMsg As String, cfg As Scripting.Dictionary
    Dim answer As VbMsgBoxResult, g As Scripting.Dictionary, existing As Scripting.Dictionary

    f = Application.GetOpenFilename("Digit presets (*.json),*.json,All files (*.*),*.*", 1, "Import Digit presets")
    If VarType(f) = vbBoolean Then Exit Sub

    If Not Store_ReadText(CStr(f), text, errMsg) Then
        MsgBox "Could not read the file:" & vbCrLf & errMsg, vbExclamation, DIGIT_TITLE
        Exit Sub
    End If
    Set cfg = Store_ParseConfig(text, errMsg)
    If cfg Is Nothing Then
        MsgBox "This is not a valid Digit presets file: " & errMsg & ".", vbExclamation, DIGIT_TITLE
        Exit Sub
    End If

    answer = MsgBox("The file has " & cfg("groups").Count & " preset group(s)." & vbCrLf & vbCrLf & _
                    "Yes: replace all presets, buttons and shortcuts with the file's." & vbCrLf & _
                    "No: add the file's groups to yours (groups with the same name are replaced)." & vbCrLf & vbCrLf & _
                    "Nothing is saved until you click Save.", _
                    vbQuestion + vbYesNoCancel, DIGIT_TITLE & " - Import")
    Select Case answer
        Case vbYes
            LoadAll cfg
        Case vbNo
            For Each g In cfg("groups")
                Set existing = Store_FindGroup(mCfg, g("name"))
                If existing Is Nothing Then
                    mCfg("groups").Add g
                Else
                    Set existing("presets") = g("presets")
                End If
            Next g
            RefreshGroups lstGroups.ListIndex + 1
            RefreshSlots
        Case Else
            Exit Sub
    End Select
    mDirty = True
End Sub

Private Sub btnExport_Click()
    Dim f As Variant, errMsg As String
    f = Application.GetSaveAsFilename("presets.json", "Digit presets (*.json),*.json", 1, "Export Digit presets")
    If VarType(f) = vbBoolean Then Exit Sub
    If Store_WriteText(CStr(f), Store_ToJson(Store_Normalize(mCfg)), errMsg) Then
        MsgBox "Presets exported to:" & vbCrLf & f & vbCrLf & vbCrLf & _
               "(This includes any unsaved changes shown in this window.)", vbInformation, DIGIT_TITLE
    Else
        MsgBox "Could not write the file:" & vbCrLf & errMsg, vbExclamation, DIGIT_TITLE
    End If
End Sub

Private Sub btnOpenFolder_Click()
    UI_OpenFolder Store_Folder()
End Sub
