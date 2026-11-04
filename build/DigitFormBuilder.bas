Attribute VB_Name = "DigitFormBuilder"
'==============================================================================
' DigitFormBuilder - creates the frmDigitSettings UserForm at build time.
'
' A UserForm's layout normally lives in a binary .frx file next to the .frm,
' which cannot be written or reviewed as plain text. Instead this module lays
' the form out from code, then loads the form's code from
' src/forms/frmDigitSettings.code.vba. build.ps1 imports this module, runs
' BuildSettingsForm, and removes the module again, so it never ships.
'
' Manual use: import this module into the add-in project, run
'   BuildSettingsForm "C:\path\to\src\forms\frmDigitSettings.code.vba"
' from the Immediate window, then delete this module.
'
' Requires "Trust access to the VBA project object model".
'==============================================================================
Option Explicit

Private Const FORM_NAME As String = "frmDigitSettings"

Public Sub BuildSettingsForm(ByVal codeFile As String)
    Dim vbp As Object, comp As Object, d As Object, c As Object, i As Long

    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents(FORM_NAME)
    On Error GoTo 0

    Set comp = vbp.VBComponents.Add(3)          ' vbext_ct_MSForm
    comp.Name = FORM_NAME
    comp.Properties("Caption").Value = "Digit - Settings"
    comp.Properties("Width").Value = 640
    comp.Properties("Height").Value = 508
    Set d = comp.Designer

    ' Two tabs: Presets (groups, pictures, formats) and Ribbon & shortcuts.
    Dim mp As Object, pg1 As Object, pg2 As Object, fr As Object, fra As Object, k As Long, y As Single
    Set mp = AddCtl(d, "Forms.MultiPage.1", "mpTabs", 4, 4, 628, 440)
    Set pg1 = mp.Pages(0): pg1.Caption = "Presets"
    Set pg2 = mp.Pages(1): pg2.Caption = "Ribbon && shortcuts"

    '=== Tab 1: Presets =======================================================
    Set fr = AddFrame(pg1, "fraGroups", 4, 4, 190, 404, "Preset groups")
    Set c = AddCtl(fr, "Forms.ListBox.1", "lstGroups", 8, 10, 172, 236)
    c.IntegralHeight = False
    Hint fr, "lblGroupsHint", 8, 250, 172, 24, "1-5 are the ribbon buttons, left to right. Move a group to change them."
    AddButton fr, "btnGroupNew", 8, 280, 54, 22, "New..."
    AddButton fr, "btnGroupRename", 64, 280, 58, 22, "Rename..."
    AddButton fr, "btnGroupDelete", 124, 280, 56, 22, "Delete"
    AddButton fr, "btnGroupUp", 8, 306, 85, 22, "Move up"
    AddButton fr, "btnGroupDown", 95, 306, 85, 22, "Move down"
    AddLabel fr, "lblIconTitle", 8, 344, 70, 12, "Group picture:"
    Set c = AddCtl(fr, "Forms.Image.1", "imgIcon", 78, 337, 26, 26)
    IconCell c
    c.PictureSizeMode = 3                         ' zoom: the 32x32 picture fits the box
    Set c = AddLabel(fr, "lblIconPick", 110, 344, 76, 12, "Choose picture...")
    c.ForeColor = &HAD5C0B                        ' link blue
    c.Font.Underline = True

    Set fr = AddFrame(pg1, "fraPresets", 200, 4, 420, 120, "Presets in this group (each click moves to the next)")
    Set c = AddCtl(fr, "Forms.ListBox.1", "lstPresets", 8, 12, 306, 74)
    c.ColumnCount = 3
    c.ColumnWidths = "14 pt;104 pt;180 pt"
    c.IntegralHeight = False
    AddButton fr, "btnPresetAdd", 320, 12, 90, 20, "Add"
    AddButton fr, "btnPresetDelete", 320, 36, 90, 20, "Remove"
    AddButton fr, "btnPresetUp", 320, 60, 44, 20, "Up"
    AddButton fr, "btnPresetDown", 366, 60, 44, 20, "Down"
    Hint fr, "lblPresetsHint", 8, 92, 400, 12, "Presets with the same number format are told apart by colour and style."

    Set fr = AddFrame(pg1, "fraEditor", 200, 130, 420, 278, "Selected preset")
    AddLabel fr, "lblFormat", 8, 16, 80, 12, "Number format"
    Set c = AddCtl(fr, "Forms.ComboBox.1", "cboFormat", 92, 13, 240, 18)
    c.Style = 0                                   ' fmStyleDropDownCombo: free text allowed
    c.MatchEntry = 2                              ' fmMatchEntryNone
    Set c = AddLabel(fr, "lblFormatStatus", 338, 16, 74, 12, "")
    c.Font.Bold = True
    AddLabel fr, "lblFont", 8, 42, 80, 12, "Font colour"
    Set c = AddLabel(fr, "lblFontSwatch", 92, 39, 40, 18, "none")
    Swatch c
    AddButton fr, "btnFontPick", 138, 38, 96, 21, "Choose colour..."
    AddButton fr, "btnFontNone", 238, 38, 56, 21, "Clear"
    AddLabel fr, "lblFill", 8, 68, 80, 12, "Fill colour"
    Set c = AddLabel(fr, "lblFillSwatch", 92, 65, 40, 18, "none")
    Swatch c
    AddButton fr, "btnFillPick", 138, 64, 96, 21, "Choose colour..."
    AddButton fr, "btnFillNone", 238, 64, 56, 21, "Clear"
    AddLabel fr, "lblStyle", 8, 94, 80, 12, "Style"
    Set c = StyleBox(fr, "chkBold", 92, 91, 64, "Bold", "B"): c.Font.Bold = True
    Set c = StyleBox(fr, "chkItalic", 160, 91, 64, "Italic", "I"): c.Font.Italic = True
    Set c = StyleBox(fr, "chkUnderline", 228, 91, 84, "Underline", "U"): c.Font.Underline = True
    Hint fr, "lblStyleHint", 92, 110, 316, 12, "Ticked = on   -   empty = off   -   grey = leave as it is"
    AddLabel fr, "lblSample", 8, 136, 80, 12, "Preview with"
    AddCtl fr, "Forms.TextBox.1", "txtSample", 92, 133, 110, 18
    Set c = AddLabel(fr, "lblPreview", 8, 158, 400, 34, "")
    c.BorderStyle = 1
    c.BackStyle = 1
    c.BackColor = &HFFFFFF
    c.Font.Size = 12
    c.TextAlign = 1
    Hint fr, "lblNone", 8, 200, 400, 24, "Colours and styles left empty (""Clear"" / grey) leave the cell as it is."

    '--- Picture picker: covers the two preset frames while open
    Set fra = AddFrame(pg1, "fraIcons", 200, 4, 420, 404, "Choose a button picture")
    fra.Visible = False
    For k = 1 To 77
        Set c = AddCtl(fra, "Forms.Image.1", "imgPick" & k, 8 + ((k - 1) Mod 11) * 36, 10 + ((k - 1) \ 11) * 36, 34, 34)
        IconCell c
    Next k
    Hint fra, "lblIconHint", 8, 268, 400, 12, "Your Digit icons come first, then Office icons. Hover an icon to see its name; click one to use it."
    AddButton fra, "btnIconPrev", 8, 286, 80, 22, "< Previous"
    Set c = AddLabel(fra, "lblIconPage", 92, 291, 100, 12, "")
    c.TextAlign = 2
    AddButton fra, "btnIconNext", 196, 286, 80, 22, "Next >"
    AddLabel fra, "lblIconName", 8, 322, 110, 12, "Or type an icon name:"
    Set c = AddCtl(fra, "Forms.TextBox.1", "txtIcon", 118, 318, 170, 18)
    c.ControlTipText = "An Office icon (imageMso) name, or digit:<name> for Digit's own. Blank = the button number."
    AddButton fra, "btnIconNone", 294, 316, 56, 22, "None"
    AddButton fra, "btnIconClose", 354, 316, 56, 22, "Close"

    '=== Tab 2: Ribbon & shortcuts ===========================================
    Set fr = AddFrame(pg2, "fraButtons", 4, 4, 616, 200, "Quick buttons on the Home tab")
    Hint fr, "lblButtonsHint", 8, 12, 600, 12, "The first five preset groups, in order. Change the order on the Presets tab."
    Set c = AddLabel(fr, "lblHdrGroup", 8, 30, 220, 12, "#    Group"): c.Font.Bold = True
    Set c = AddLabel(fr, "lblHdrKey", 250, 30, 160, 12, "Keyboard shortcut"): c.Font.Bold = True
    For i = 1 To 5
        y = 46 + (i - 1) * 24
        AddLabel fr, "lblSlot" & i, 8, y + 3, 236, 12, i & ":"
        Set c = AddCtl(fr, "Forms.TextBox.1", "txtKey" & i, 250, y, 140, 18)
        c.ControlTipText = "For example Ctrl+Alt+" & i & ". Leave blank for no shortcut."
    Next i
    Set c = AddCtl(fr, "Forms.CheckBox.1", "chkShortcuts", 8, 172, 220, 16)
    c.Caption = "Enable keyboard shortcuts"

    Set fr = AddFrame(pg2, "fraRibbon", 4, 210, 616, 58, "Ribbon")
    Set c = AddCtl(fr, "Forms.CheckBox.1", "chkLarge", 8, 12, 300, 16)
    c.Caption = "Large quick buttons (one per column)"
    Hint fr, "lblRibbonHint", 8, 32, 600, 12, "Small: three to a column, Undo Format bottom right. Large: one per column."

    Set fr = AddFrame(pg2, "fraUpdates", 4, 274, 616, 40, "Updates")
    Set c = AddCtl(fr, "Forms.CheckBox.1", "chkUpdates", 8, 12, 400, 16)
    c.Caption = "Check for updates once a day when Excel starts"

    Set fr = AddFrame(pg2, "fraFile", 4, 320, 616, 84, "Presets file")
    Hint fr, "lblFileHint", 8, 12, 600, 12, "Share presets with your team, or keep a backup."
    AddButton fr, "btnImport", 8, 32, 80, 24, "Import..."
    AddButton fr, "btnExport", 92, 32, 80, 24, "Export..."
    AddButton fr, "btnOpenFolder", 176, 32, 90, 24, "Open folder"

    '=== Footer (both tabs) ===================================================
    Set c = Hint(d, "lblStatus", 8, 456, 460, 12, "")
    Set c = AddButton(d, "btnSave", 478, 450, 72, 24, "Save")
    c.Default = True
    c.Accelerator = "S"
    Set c = AddButton(d, "btnCancel", 556, 450, 72, 24, "Cancel")
    c.Cancel = True

    '--- Code
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromFile codeFile
    End With
End Sub

Private Function AddCtl(ByVal d As Object, ByVal progId As String, ByVal nm As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = d.Controls.Add(progId, nm, True)
    c.Left = l
    c.Top = t
    c.Width = w
    c.Height = h
    Set AddCtl = c
End Function

Private Function AddLabel(ByVal d As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim c As Object
    Set c = AddCtl(d, "Forms.Label.1", nm, l, t, w, h)
    c.Caption = cap
    c.WordWrap = True
    Set AddLabel = c
End Function

Private Function AddButton(ByVal d As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                           ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim c As Object
    Set c = AddCtl(d, "Forms.CommandButton.1", nm, l, t, w, h)
    c.Caption = cap
    Set AddButton = c
End Function

Private Function AddFrame(ByVal parent As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim c As Object
    Set c = AddCtl(parent, "Forms.Frame.1", nm, l, t, w, h)
    c.Caption = cap
    Set AddFrame = c
End Function

' Small explanatory text in dark grey (about 6:1 contrast on the form background).
Private Function Hint(ByVal parent As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                      ByVal w As Single, ByVal h As Single, ByVal cap As String) As Object
    Dim c As Object
    Set c = AddLabel(parent, nm, l, t, w, h, cap)
    c.ForeColor = &H5A5A5A
    Set Hint = c
End Function

' Three-state check box: ticked = on, empty = off, grey (Null) = leave as is.
Private Function StyleBox(ByVal parent As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal cap As String, ByVal accel As String) As Object
    Dim c As Object
    Set c = AddCtl(parent, "Forms.CheckBox.1", nm, l, t, w, 16)
    c.Caption = cap
    c.TripleState = True
    c.Accelerator = accel
    c.ControlTipText = "Ticked: make it " & LCase$(cap) & ". Empty: remove " & LCase$(cap) & ". Grey: leave as it is."
    Set StyleBox = c
End Function

Private Sub IconCell(ByVal c As Object)
    c.BorderStyle = 0
    c.SpecialEffect = 0
    c.BackStyle = 1
    c.BackColor = &HFFFFFF
    c.PictureSizeMode = 0                         ' draw at its own size
    c.PictureAlignment = 2                        ' centre
End Sub

Private Sub Swatch(ByVal c As Object)
    c.BorderStyle = 1                             ' fmBorderStyleSingle
    c.BackStyle = 1                               ' fmBackStyleOpaque
    c.TextAlign = 2                               ' centre
    c.WordWrap = False
End Sub
