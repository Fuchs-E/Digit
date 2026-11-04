Attribute VB_Name = "DigitRibbon"
'==============================================================================
' DigitRibbon - callbacks for customUI14.xml.
'
' Thin layer: reads labels and menu content from the presets and forwards clicks
' to DigitApp. No formatting logic here.
'
' The IRibbonUI handle is lost if VBA is reset (an unhandled error, or pressing
' Stop in the editor). Its pointer is kept on the add-in's hidden sheet so
' Invalidate keeps working afterwards.
'==============================================================================
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

Private Const RIBBON_PTR_CELL As String = "B1"
Private Const MENU_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"

Private mRibbon As IRibbonUI

'------------------------------------------------------------------------------
' Ribbon handle
'------------------------------------------------------------------------------

Public Sub Digit_RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Dim wasSaved As Boolean
    Set mRibbon = ribbon
    On Error Resume Next
    wasSaved = ThisWorkbook.Saved
    ThisWorkbook.Worksheets(1).Range(RIBBON_PTR_CELL).Value = "'" & CStr(ObjPtr(ribbon))
    ThisWorkbook.Saved = wasSaved
    Digit_Startup
End Sub

' Redraws every Digit control (labels, visibility, menu, undo state).
Public Sub Ribbon_Invalidate()
    Dim r As IRibbonUI
    Set r = GetRibbon()
    If Not r Is Nothing Then r.Invalidate
End Sub

Public Sub Ribbon_InvalidateControl(ByVal controlId As String)
    Dim r As IRibbonUI
    Set r = GetRibbon()
    If Not r Is Nothing Then r.InvalidateControl controlId
End Sub

Private Function GetRibbon() As IRibbonUI
    Dim stored As String, p As LongPtr, obj As Object, zero As LongPtr
    If Not mRibbon Is Nothing Then
        Set GetRibbon = mRibbon
        Exit Function
    End If
    On Error GoTo Done
    stored = CStr(ThisWorkbook.Worksheets(1).Range(RIBBON_PTR_CELL).Value)
    If Len(stored) = 0 Or Not IsNumeric(stored) Then Exit Function
    p = CLngPtr(stored)
    If p = 0 Then Exit Function
    CopyMemory obj, p, LenB(p)
    Set mRibbon = obj
    CopyMemory obj, zero, LenB(p)   ' drop the borrowed reference without releasing it
    Set GetRibbon = mRibbon
Done:
End Function

'------------------------------------------------------------------------------
' Quick buttons (tag = slot number)
'------------------------------------------------------------------------------

Public Sub Digit_OnQuick(ByVal control As IRibbonControl)
    App_RunSlot SlotOf(control)
End Sub

Public Sub Digit_GetLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim nm As String
    nm = Store_SlotGroupName(Store_Config(), SlotOf(control))
    If Len(nm) = 0 Then nm = "Button " & SlotOf(control)
    returnedVal = Core_RibbonLabel(nm)
End Sub

' Slots without a group are hidden; groups not on a button are in "More presets".
Public Sub Digit_GetVisible(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Len(Store_SlotGroupName(Store_Config(), SlotOf(control))) > 0)
End Sub

Public Sub Digit_GetScreentip(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim nm As String, key As String
    nm = Store_SlotGroupName(Store_Config(), SlotOf(control))
    key = App_ActiveShortcut(SlotOf(control))
    If Len(key) > 0 Then nm = nm & " (" & key & ")"
    returnedVal = nm
End Sub

Public Sub Digit_GetSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = GroupTip(Store_FindGroup(Store_Config(), Store_SlotGroupName(Store_Config(), SlotOf(control))))
End Sub

' Picture for a quick button: the group's "icon" when it names an Office icon,
' otherwise the button number. Returning a name (not a picture) makes the
' ribbon draw that built-in icon.
Public Sub Digit_GetImage(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = ButtonIcon(Store_FindGroup(Store_Config(), Store_SlotGroupName(Store_Config(), SlotOf(control))), _
                             "_" & SlotOf(control))
End Sub

Private Function ButtonIcon(ByVal g As Scripting.Dictionary, ByVal fallback As String) As String
    Dim nm As String
    nm = Store_GroupIcon(g)
    If UI_IsOfficeIcon(nm) Then ButtonIcon = nm Else ButtonIcon = fallback
End Function

Private Function SlotOf(ByVal control As IRibbonControl) As Long
    If IsNumeric(control.Tag) Then SlotOf = CLng(control.Tag)
End Function

'------------------------------------------------------------------------------
' "More presets" dynamic menu (button tag = group name)
'------------------------------------------------------------------------------

Public Sub Digit_GetMoreContent(ByVal control As IRibbonControl, ByRef returnedVal)
    Dim cfg As Scripting.Dictionary, g As Scripting.Dictionary, xml As String, i As Long, n As Long

    Set cfg = Store_Config()
    xml = "<menu xmlns=""" & MENU_NS & """>"
    For Each g In cfg("groups")
        i = i + 1
        If Not Store_IsInSlot(cfg, g("name")) Then
            n = n + 1
            xml = xml & "<button id=""dgMore" & i & """" & _
                  " label=""" & Core_XmlEscape(Core_RibbonLabel(g("name"))) & """" & _
                  " tag=""" & Core_XmlEscape(g("name")) & """" & _
                  MenuIcon(g) & _
                  " supertip=""" & Core_XmlEscape(GroupTip(g)) & """" & _
                  " onAction=""Digit_OnMoreItem""/>"
        End If
    Next g
    If n = 0 Then
        xml = xml & "<button id=""dgMoreNone"" label=""(every group is on a button)"" enabled=""false""/>"
    End If
    xml = xml & "<menuSeparator id=""dgMoreSep""/>" & _
          "<button id=""dgMoreSettings"" label=""Manage presets..."" imageMso=""PropertySheet"" onAction=""Digit_OnSettings""/>" & _
          "</menu>"
    returnedVal = xml
End Sub

Private Function MenuIcon(ByVal g As Scripting.Dictionary) As String
    Dim nm As String
    nm = ButtonIcon(g, "")
    If Len(nm) > 0 Then MenuIcon = " imageMso=""" & Core_XmlEscape(nm) & """"
End Function

Public Sub Digit_OnMoreItem(ByVal control As IRibbonControl)
    App_RunGroup control.Tag
End Sub

'------------------------------------------------------------------------------
' Tools
'------------------------------------------------------------------------------

Public Sub Digit_OnUndo(ByVal control As IRibbonControl)
    Digit_UndoLast
End Sub

Public Sub Digit_GetUndoEnabled(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Undo_Count() > 0)
End Sub

Public Sub Digit_GetUndoSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
    If Undo_Count() > 0 Then
        returnedVal = "Undo the last Digit change: " & Undo_PeekLabel() & "." & vbLf & _
                      Undo_Count() & " change(s) can be undone. Ctrl+Z also works straight after a change."
    Else
        returnedVal = "Nothing to undo yet. Digit keeps its own history of the last formatting changes."
    End If
End Sub

Public Sub Digit_OnSettings(ByVal control As IRibbonControl)
    App_ShowSettings
End Sub

Public Sub Digit_OnReload(ByVal control As IRibbonControl)
    App_ReloadPresets
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------

' Super-tip text listing a group's presets in cycle order.
Private Function GroupTip(ByVal g As Scripting.Dictionary) As String
    Dim p As Scripting.Dictionary, s As String, i As Long
    If g Is Nothing Then Exit Function
    s = "Click to cycle through:"
    For Each p In g("presets")
        i = i + 1
        s = s & vbLf & i & ". " & p("numberFormat")
        If Len(p("fontColor")) > 0 Then s = s & "  font " & p("fontColor")
        If Len(p("fillColor")) > 0 Then s = s & "  fill " & p("fillColor")
    Next p
    If Len(s) > 1000 Then s = Left$(s, 997) & "..."
    GroupTip = s
End Function
