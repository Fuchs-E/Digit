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
    Dim vbp As Object, comp As Object, d As Object, c As Object, i As Long, x As Single

    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents(FORM_NAME)
    On Error GoTo 0

    Set comp = vbp.VBComponents.Add(3)          ' vbext_ct_MSForm
    comp.Name = FORM_NAME
    comp.Properties("Caption").Value = "Digit - Presets"
    comp.Properties("Width").Value = 640
    comp.Properties("Height").Value = 458
    Set d = comp.Designer

    '--- Groups (left)
    AddLabel d, "lblGroups", 8, 8, 180, 12, "Preset groups"
    Set c = AddCtl(d, "Forms.ListBox.1", "lstGroups", 8, 22, 180, 200)
    c.IntegralHeight = False
    AddButton d, "btnGroupNew", 8, 228, 58, 22, "New..."
    AddButton d, "btnGroupRename", 69, 228, 58, 22, "Rename..."
    AddButton d, "btnGroupDelete", 130, 228, 58, 22, "Delete"
    AddButton d, "btnGroupUp", 8, 254, 88, 22, "Move up"
    AddButton d, "btnGroupDown", 100, 254, 88, 22, "Move down"

    '--- Presets of the selected group (right)
    AddLabel d, "lblPresets", 200, 8, 330, 12, "Presets in cycle order (up to 3)"
    Set c = AddCtl(d, "Forms.ListBox.1", "lstPresets", 200, 22, 330, 62)
    c.ColumnCount = 3
    c.ColumnWidths = "16 pt;170 pt;130 pt"
    c.IntegralHeight = False
    AddButton d, "btnPresetAdd", 536, 22, 92, 20, "Add preset"
    AddButton d, "btnPresetDelete", 536, 44, 92, 20, "Remove"
    AddButton d, "btnPresetUp", 536, 66, 45, 20, "Up"
    AddButton d, "btnPresetDown", 583, 66, 45, 20, "Down"

    AddLabel d, "lblFormat", 200, 96, 200, 12, "Number format (pick or type)"
    Set c = AddCtl(d, "Forms.ComboBox.1", "cboFormat", 200, 110, 330, 18)
    c.Style = 0                                   ' fmStyleDropDownCombo: free text allowed
    c.MatchEntry = 2                              ' fmMatchEntryNone: do not auto-complete while typing
    Set c = AddLabel(d, "lblFormatStatus", 536, 113, 92, 12, "")
    c.Font.Bold = True

    AddLabel d, "lblFont", 200, 137, 70, 12, "Font colour"
    Set c = AddLabel(d, "lblFontSwatch", 272, 134, 44, 17, "none")
    Swatch c
    AddButton d, "btnFontPick", 322, 133, 64, 19, "Choose..."
    AddButton d, "btnFontNone", 390, 133, 64, 19, "None"

    AddLabel d, "lblFill", 200, 160, 70, 12, "Fill colour"
    Set c = AddLabel(d, "lblFillSwatch", 272, 157, 44, 17, "none")
    Swatch c
    AddButton d, "btnFillPick", 322, 156, 64, 19, "Choose..."
    AddButton d, "btnFillNone", 390, 156, 64, 19, "None"

    AddLabel d, "lblNone", 460, 137, 168, 36, """None"" leaves the cell's existing colour as it is."

    AddLabel d, "lblSample", 200, 184, 70, 12, "Sample value"
    AddCtl d, "Forms.TextBox.1", "txtSample", 272, 181, 114, 18

    AddLabel d, "lblPreviewCaption", 200, 206, 300, 12, "Preview: positive, negative, zero"
    Set c = AddLabel(d, "lblPreview", 200, 220, 428, 28, "")
    c.BorderStyle = 1
    c.BackStyle = 1
    c.BackColor = &HFFFFFF
    c.Font.Size = 12
    c.TextAlign = 1                               ' left

    '--- Quick buttons and shortcuts (bottom)
    AddLabel d, "lblSlots", 8, 288, 620, 12, "Quick buttons on the ribbon (preset group, then keyboard shortcut)"
    For i = 1 To 5
        x = 8 + (i - 1) * 125
        AddLabel d, "lblSlot" & i, x, 304, 120, 12, "Button " & i
        Set c = AddCtl(d, "Forms.ComboBox.1", "cboSlot" & i, x, 318, 120, 18)
        c.Style = 2                               ' fmStyleDropDownList
        Set c = AddCtl(d, "Forms.TextBox.1", "txtKey" & i, x, 340, 120, 18)
        c.ControlTipText = "For example Ctrl+Alt+" & i & ". Leave blank for no shortcut."
    Next i
    Set c = AddCtl(d, "Forms.CheckBox.1", "chkShortcuts", 8, 364, 160, 16)
    c.Caption = "Enable keyboard shortcuts"
    AddLabel d, "lblHint", 172, 366, 456, 12, ""

    '--- Bottom row
    AddButton d, "btnImport", 8, 398, 70, 24, "Import..."
    AddButton d, "btnExport", 82, 398, 70, 24, "Export..."
    AddButton d, "btnOpenFolder", 156, 398, 76, 24, "Open folder"
    Set c = AddLabel(d, "lblStatus", 238, 404, 230, 12, "")
    c.ForeColor = &H808080
    Set c = AddButton(d, "btnSave", 478, 398, 72, 24, "Save")
    c.Default = True
    Set c = AddButton(d, "btnCancel", 556, 398, 72, 24, "Cancel")
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

Private Sub Swatch(ByVal c As Object)
    c.BorderStyle = 1                             ' fmBorderStyleSingle
    c.BackStyle = 1                               ' fmBackStyleOpaque
    c.TextAlign = 2                               ' centre
    c.WordWrap = False
End Sub
