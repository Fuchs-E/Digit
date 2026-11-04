Attribute VB_Name = "DigitApp"
'==============================================================================
' DigitApp - glue between the ribbon / keyboard / settings window and the
' engine. Owns start-up and shut-down, keyboard shortcuts, user messages and
' Ctrl+Z wiring.
'==============================================================================
Option Explicit

Public Const DIGIT_VERSION As String = "1.1.5"
Public Const DIGIT_TITLE As String = "Digit"

Private mStarted As Boolean
Private mKeys(1 To DIGIT_SLOT_COUNT) As String   ' OnKey codes currently registered
Private mShortcutLabels(1 To DIGIT_SLOT_COUNT) As String
Private mStatusText As String
Private mStatusClearAt As Date

'------------------------------------------------------------------------------
' Life cycle (called from ThisWorkbook events and the ribbon onLoad)
'------------------------------------------------------------------------------

Public Sub Digit_Startup()
    Dim note As String
    If mStarted Then Exit Sub
    mStarted = True
    On Error Resume Next
    Store_Load
    note = Store_TakeLoadMessage()
    Undo_SetDepth Store_UndoDepth(Store_Config())
    App_RegisterShortcuts
    On Error GoTo 0
    If Len(note) > 0 Then
        ' A hidden Excel (automation, build tests) must never block on a dialog.
        If Application.Visible Then MsgBox note, vbExclamation, DIGIT_TITLE Else Debug.Print note
    End If
    ScheduleAutoUpdateCheck
End Sub

Public Sub Digit_Shutdown()
    On Error Resume Next
    App_UnregisterShortcuts
    If mStatusClearAt > 0 Then Application.OnTime mStatusClearAt, Qualified("Digit_ClearStatus"), , False
    If Len(mStatusText) > 0 Then
        If Application.StatusBar = mStatusText Then Application.StatusBar = False
    End If
    mStarted = False
End Sub

' Re-runs everything that depends on the presets (after Settings > Save,
' Reload, or an import).
Public Sub App_OnPresetsChanged()
    On Error Resume Next
    Fmt_ClearCache
    Undo_SetDepth Store_UndoDepth(Store_Config())
    App_RegisterShortcuts
    Ribbon_Invalidate
End Sub

Public Sub App_ReloadPresets()
    Dim note As String
    Store_Load
    note = Store_TakeLoadMessage()
    App_OnPresetsChanged
    If Len(note) > 0 Then
        MsgBox note, vbExclamation, DIGIT_TITLE
    Else
        App_SetStatus "Digit: presets reloaded from " & Store_FilePath()
    End If
End Sub

'------------------------------------------------------------------------------
' Running presets
'------------------------------------------------------------------------------

Public Sub App_RunSlot(ByVal slot As Long)
    Dim nm As String
    Digit_Startup
    nm = Store_SlotGroupName(Store_Config(), slot)
    If Len(nm) = 0 Then
        If MsgBox("Quick button " & slot & " has no preset group yet." & vbCrLf & vbCrLf & _
                  "Open Digit Settings to choose one?", vbQuestion + vbYesNo, DIGIT_TITLE) = vbYes Then
            App_ShowSettings
        End If
        Exit Sub
    End If
    App_RunGroup nm
End Sub

Public Sub App_RunGroup(ByVal groupName As String)
    Dim grp As Scripting.Dictionary, res As DigitResult, msg As String

    Digit_Startup
    Set grp = Store_FindGroup(Store_Config(), groupName)
    res = Fmt_ApplyGroup(grp, msg)

    If res = drUndoTooBig Then
        If MsgBox(msg & vbCrLf & vbCrLf & "Apply the format anyway? It cannot be undone.", _
                  vbExclamation + vbYesNo + vbDefaultButton2, DIGIT_TITLE) <> vbYes Then Exit Sub
        res = Fmt_ApplyGroup(grp, msg, True)
    End If

    App_Report res, msg
    If res = drOK Then App_AfterChange
End Sub

' After a change: hook Ctrl+Z to Digit's undo and refresh the Undo button.
Private Sub App_AfterChange()
    On Error Resume Next
    If Undo_Count() > 0 Then
        Application.OnUndo "Undo Digit format (" & Undo_PeekLabel() & ")", Qualified("Digit_UndoLast")
    End If
    Ribbon_InvalidateControl "btnDigitUndo"
End Sub

' Undo Format button, and Ctrl+Z straight after a Digit change.
Public Sub Digit_UndoLast()
    Dim res As DigitResult, msg As String
    res = Undo_Last(msg)
    App_Report res, msg
    On Error Resume Next
    ' Ctrl+Z only reaches Digit straight after a Digit change, so re-arm it for
    ' the next step back while history remains.
    If Undo_Count() > 0 Then
        Application.OnUndo "Undo Digit format (" & Undo_PeekLabel() & ")", Qualified("Digit_UndoLast")
    End If
    Ribbon_InvalidateControl "btnDigitUndo"
End Sub

' Targets for Application.OnKey (they must take no arguments).
Public Sub DigitQuick1(): App_RunSlot 1: End Sub
Public Sub DigitQuick2(): App_RunSlot 2: End Sub
Public Sub DigitQuick3(): App_RunSlot 3: End Sub
Public Sub DigitQuick4(): App_RunSlot 4: End Sub
Public Sub DigitQuick5(): App_RunSlot 5: End Sub

'------------------------------------------------------------------------------
' Keyboard shortcuts
'------------------------------------------------------------------------------

Public Sub App_RegisterShortcuts()
    Dim cfg As Scripting.Dictionary, i As Long, code As String

    App_UnregisterShortcuts
    Set cfg = Store_Config()
    If Not Store_ShortcutsEnabled(cfg) Then Exit Sub

    For i = 1 To DIGIT_SLOT_COUNT
        code = Core_ShortcutToOnKey(Store_Shortcut(cfg, i))
        If Len(code) > 0 And Len(Store_SlotGroupName(cfg, i)) > 0 Then
            On Error Resume Next
            Err.Clear
            Application.OnKey code, Qualified("DigitQuick" & i)
            If Err.Number = 0 Then
                mKeys(i) = code
                mShortcutLabels(i) = Store_Shortcut(cfg, i)
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Public Sub App_UnregisterShortcuts()
    Dim i As Long
    On Error Resume Next
    For i = 1 To DIGIT_SLOT_COUNT
        If Len(mKeys(i)) > 0 Then Application.OnKey mKeys(i)   ' no procedure = restore Excel's default
        mKeys(i) = ""
        mShortcutLabels(i) = ""
    Next i
End Sub

' Shortcut text for a slot if it is currently registered, else "".
Public Function App_ActiveShortcut(ByVal slot As Long) As String
    If slot < 1 Or slot > DIGIT_SLOT_COUNT Then Exit Function
    App_ActiveShortcut = mShortcutLabels(slot)
End Function

'------------------------------------------------------------------------------
' Settings window
'------------------------------------------------------------------------------

Public Sub App_ShowSettings()
    Dim f As frmDigitSettings
    Digit_Startup
    Set f = New frmDigitSettings
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub

'------------------------------------------------------------------------------
' Messages
'------------------------------------------------------------------------------

Public Sub App_Report(ByVal res As DigitResult, ByVal msg As String)
    If Len(msg) = 0 Then Exit Sub
    Select Case res
        Case drOK, drNothingToDo
            App_SetStatus msg
        Case drWarning, drUndoTooBig
            MsgBox msg, vbExclamation, DIGIT_TITLE
        Case Else
            MsgBox msg, vbCritical, DIGIT_TITLE
    End Select
End Sub

' Shows text on the status bar and clears it a few seconds later.
Public Sub App_SetStatus(ByVal text As String)
    On Error Resume Next
    If mStatusClearAt > 0 Then Application.OnTime mStatusClearAt, Qualified("Digit_ClearStatus"), , False
    mStatusText = text
    Application.StatusBar = text
    mStatusClearAt = Now + TimeSerial(0, 0, 5)
    Application.OnTime mStatusClearAt, Qualified("Digit_ClearStatus")
End Sub

Public Sub Digit_ClearStatus()
    On Error Resume Next
    mStatusClearAt = 0
    ' Leave the status bar alone if something else has written to it since.
    If Application.StatusBar = mStatusText Then Application.StatusBar = False
    mStatusText = ""
End Sub

'------------------------------------------------------------------------------
' Updates
'------------------------------------------------------------------------------

' Once a day, shortly after Excel starts, look for a new release. Never in a
' hidden Excel (builds, tests) and never when turned off in presets.json.
Private Sub ScheduleAutoUpdateCheck()
    On Error Resume Next
    If Not Application.Visible Then Exit Sub
    If Not Store_CheckForUpdates(Store_Config()) Then Exit Sub
    If Update_GetSetting("LastCheck") = Format$(Date, "yyyy-mm-dd") Then Exit Sub
    Application.OnTime Now + TimeSerial(0, 0, 20), Qualified("Digit_AutoUpdateCheck")
End Sub

Public Sub Digit_AutoUpdateCheck()
    App_CheckForUpdates True
End Sub

' silent = True: automatic daily check. Says nothing unless there is a new
' version, and does not offer a version the user chose to skip.
Public Sub App_CheckForUpdates(ByVal silent As Boolean)
    Dim tag As String, notes As String, pageUrl As String, errMsg As String
    Dim prompt As String, answer As VbMsgBoxResult

    Digit_Startup
    Update_SaveSetting "LastCheck", Format$(Date, "yyyy-mm-dd")
    If Not silent Then App_SetStatus "Digit: checking for updates..."

    If Not Update_FetchLatest(tag, notes, pageUrl, errMsg) Then
        If Not silent Then MsgBox "Digit could not check for updates." & vbCrLf & vbCrLf & errMsg, vbExclamation, DIGIT_TITLE
        Exit Sub
    End If

    If Core_CompareVersions(tag, DIGIT_VERSION) <= 0 Then
        If Not silent Then MsgBox "You have the latest version of Digit (" & DIGIT_VERSION & ").", vbInformation, DIGIT_TITLE
        Exit Sub
    End If

    If Update_GetSetting("Pending") = tag Then
        If Not silent Then
            MsgBox "Digit " & tag & " is already being installed." & vbCrLf & vbCrLf & _
                   "Close Excel to finish. If you already did and it did not install, see:" & vbCrLf & _
                   Store_Folder() & "\update\update.log", vbInformation, DIGIT_TITLE
        End If
        Exit Sub
    End If
    If silent And Update_GetSetting("Skip") = tag Then Exit Sub

    prompt = "Digit " & tag & " is available. You have " & DIGIT_VERSION & "." & vbCrLf & vbCrLf
    If Len(Trim$(notes)) > 0 Then prompt = prompt & "What's new:" & vbCrLf & Left$(Trim$(notes), 600) & vbCrLf & vbCrLf
    prompt = prompt & "Install it now? Digit builds and tests it in the background (about a minute), " & _
                      "then installs it when you close Excel. You can keep working."
    If silent Then
        prompt = prompt & vbCrLf & vbCrLf & "Yes = install    No = remind me tomorrow    Cancel = skip this version"
        answer = MsgBox(prompt, vbQuestion + vbYesNoCancel, DIGIT_TITLE)
        If answer = vbCancel Then Update_SaveSetting "Skip", tag
    Else
        answer = MsgBox(prompt, vbQuestion + vbYesNo, DIGIT_TITLE)
    End If
    If answer <> vbYes Then Exit Sub

    App_SetStatus "Digit: downloading " & tag & "..."
    If Not Update_Start(tag, errMsg) Then
        MsgBox "Digit could not start the update." & vbCrLf & vbCrLf & errMsg, vbExclamation, DIGIT_TITLE
        Exit Sub
    End If
    Update_SaveSetting "Pending", tag
    MsgBox "Digit " & tag & " is being built and tested in the background." & vbCrLf & vbCrLf & _
           "Keep working. When you next close Excel, it is installed, and a message confirms it." & vbCrLf & _
           "If anything fails, your current version stays as it is.", vbInformation, DIGIT_TITLE
End Sub

' "'Digit.xlam'!Proc" so OnKey / OnUndo / OnTime always find the add-in's macro.
Private Function Qualified(ByVal procName As String) As String
    Qualified = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function
