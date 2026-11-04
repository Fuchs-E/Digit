Attribute VB_Name = "DigitUI"
'==============================================================================
' DigitUI - small Windows helpers for the settings window: the standard colour
' picker and opening a folder in Explorer. 64-bit and 32-bit Office (VBA7).
'==============================================================================
Option Explicit

Private Type DIGIT_CHOOSECOLOR
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    Flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type

Private Declare PtrSafe Function ChooseColorA Lib "comdlg32.dll" (ByRef lpcc As DIGIT_CHOOSECOLOR) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_ANYCOLOR As Long = &H100

Private mCustomColors(0 To 15) As Long

' Icons offered in the settings window. Only names this Office version knows
' are shown, so an unknown name here is harmless.
Public Const DIGIT_ICON_SUGGESTIONS As String = _
    "_1,_2,_3,_4,_5,_6,_7,_8,_9,PercentStyle,CommaStyle,DecimalsIncrease,DecimalsDecrease," & _
    "AccountingFormat,CurrencySymbol,NumberFormatsDialog,FormatNumberDefault,FormatCellsDialog," & _
    "FontColorPicker,CellFillColorPicker,FormatPainter,DateAndTimeInsert,CalendarInsert," & _
    "Calculator,ChartTypeOtherInsertGallery,TableInsert,HappyFace,StarRatingFull,FlagRed," & _
    "A,B,C,D,E,F,G,H,K,M,N,P,T,X"
Private mCustomInit As Boolean

' Shows the Windows colour dialog. Returns the chosen colour (BGR Long, as used
' by Excel and MSForms), or -1 when cancelled.
'   initialColor  -1 for none
'   ownerCaption  caption of the UserForm that owns the dialog
Public Function UI_PickColor(ByVal initialColor As Long, ByVal ownerCaption As String) As Long
    Dim cc As DIGIT_CHOOSECOLOR, i As Long

    If Not mCustomInit Then
        ' Seed the custom colours with common finance-model colours.
        mCustomColors(0) = RGB(0, 0, 255)       ' input blue
        mCustomColors(1) = RGB(0, 128, 0)       ' link green
        mCustomColors(2) = RGB(255, 0, 0)       ' red
        mCustomColors(3) = RGB(255, 242, 204)   ' light yellow
        mCustomColors(4) = RGB(226, 239, 218)   ' light green
        mCustomColors(5) = RGB(221, 235, 247)   ' light blue
        mCustomColors(6) = RGB(252, 228, 214)   ' light orange
        mCustomColors(7) = RGB(242, 242, 242)   ' light grey
        For i = 8 To 15
            mCustomColors(i) = RGB(255, 255, 255)
        Next i
        mCustomInit = True
    End If

    cc.lStructSize = LenB(cc)
    cc.hwndOwner = FindWindowA("ThunderDFrame", ownerCaption)
    cc.lpCustColors = VarPtr(mCustomColors(0))
    cc.Flags = CC_ANYCOLOR Or CC_FULLOPEN
    If initialColor >= 0 Then
        cc.rgbResult = initialColor
        cc.Flags = cc.Flags Or CC_RGBINIT
    End If

    If ChooseColorA(cc) <> 0 Then
        UI_PickColor = cc.rgbResult
    Else
        UI_PickColor = -1
    End If
End Function

Public Sub UI_OpenFolder(ByVal folder As String)
    On Error Resume Next
    Shell "explorer.exe """ & folder & """", vbNormalFocus
End Sub

'------------------------------------------------------------------------------
' Office icons (imageMso names)
'------------------------------------------------------------------------------

' True when name is a built-in Office icon this version of Excel can draw.
Public Function UI_IsOfficeIcon(ByVal name As String) As Boolean
    Dim pic As Object
    If Len(Trim$(name)) = 0 Then Exit Function
    On Error Resume Next
    Set pic = Application.CommandBars.GetImageMso(name, 16, 16)
    UI_IsOfficeIcon = (Err.Number = 0 And Not pic Is Nothing)
End Function

' The icon as a picture (32x32), or Nothing when the name is unknown.
Public Function UI_OfficeIconPicture(ByVal name As String) As Object
    If Len(Trim$(name)) = 0 Then Exit Function
    On Error Resume Next
    Set UI_OfficeIconPicture = Application.CommandBars.GetImageMso(name, 32, 32)
End Function
