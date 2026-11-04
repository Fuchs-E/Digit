'==============================================================================
' Code for the add-in's ThisWorkbook module.
'
' This file is not imported as a module: build.ps1 pastes its text into the
' workbook's document module (ThisWorkbook, whatever its localised name).
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    Digit_Startup
End Sub

Private Sub Workbook_AddinInstall()
    Digit_Startup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Digit_Shutdown
End Sub

Private Sub Workbook_AddinUninstall()
    Digit_Shutdown
End Sub
