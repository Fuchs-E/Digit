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
' Pictures for quick buttons: Digit's own icons ("digit:...") and built-in
' Office icons (imageMso names).
'------------------------------------------------------------------------------

' Every picture offered in the settings window, Digit's own first. Office icons
' this version of Excel does not have are left out, so the list below can be
' generous.
Public Function UI_IconCandidates() As Collection
    Dim c As New Collection, nm As Variant, s As String
    For Each nm In Icons_List()
        c.Add CStr(nm)
    Next nm
    s = s & "_1,_2,_3,_4,_5,_6,_7,_8,_9,_0,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,PercentStyle,CommaStyle,"
    s = s & "AccountingFormat,AccountingFormatMenu,CurrencySymbol,NumberFormatGallery,NumberFormatsDialog,DecimalsIncrease,"
    s = s & "DecimalsDecrease,FormatNumberDefault,FormatCellsDialog,DateAndTimeInsert,CalendarInsert,DateTimeInsert,"
    s = s & "AutoSum,AutoSumMenu,FunctionWizard,Calculator,FormulaMoreFunctionsMenu,FormulaFinancialInsertGallery,"
    s = s & "FormulaLogicalInsertGallery,FormulaTextInsertGallery,FormulaDateTimeInsertGallery,FormulaLookupReferenceInsertGallery,"
    s = s & "FormulaMathTrigInsertGallery,NameManager,NameDefine,FormulaEvaluate,TraceDependents,TracePrecedents,"
    s = s & "ShowFormulas,CalculateNow,CalculateSheet,SymbolInsert,EquationInsertNew,EquationProfessional,FontColorPicker,"
    s = s & "CellFillColorPicker,FillColor,Bold,Italic,Underline,Strikethrough,FontSizeIncrease,FontSizeDecrease,"
    s = s & "Superscript,Subscript,ChangeCase,ClearFormats,Clear,ClearAll,ClearMenu,ClearContents,BordersMenu,"
    s = s & "BorderBottom,BorderTop,BorderLeft,BorderRight,BorderAll,BorderOutside,BorderThickBox,BorderDoubleBottom,"
    s = s & "BorderNone,AlignLeft,AlignCenter,AlignRight,AlignTop,AlignMiddle,AlignBottom,AlignJustify,IndentIncrease,"
    s = s & "IndentDecrease,WrapText,MergeCenter,MergeCellsMenu,MergeCells,UnmergeCells,OrientationMenu,TextDirectionMenu,"
    s = s & "ConditionalFormattingMenu,ConditionalFormattingHighlightCellsMenu,ConditionalFormattingTopBottomMenu,"
    s = s & "ConditionalFormattingDataBarsGallery,ConditionalFormattingColorScalesGallery,ConditionalFormattingIconSetsGallery,"
    s = s & "ConditionalFormattingClearMenu,CellStylesGallery,FormatPainter,FormatAsTableGallery,TableStyleClear,"
    s = s & "TableInsert,PivotTableInsert,PivotChartInsert,PivotTableFieldList,PivotTableRefresh,RefreshAll,Refresh,"
    s = s & "ChartInsert,ChartTypeColumnInsertGallery,ChartTypeLineInsertGallery,ChartTypePieInsertGallery,ChartTypeBarInsertGallery,"
    s = s & "ChartTypeAreaInsertGallery,ChartTypeXYScatterInsertGallery,ChartTypeOtherInsertGallery,SparklineLineInsert,"
    s = s & "SparklineColumnInsert,SparklineWinLossInsert,SortAscendingExcel,SortDescendingExcel,SortCustomExcel,"
    s = s & "Filter,FilterClearAllFilters,FilterReapply,AdvancedFilterDialog,GroupSort,Group,Ungroup,OutlineSubtotals,"
    s = s & "OutlineGroup,OutlineUngroup,OutlineShowDetail,OutlineHideDetail,DataValidation,DataValidationMenu,"
    s = s & "RemoveDuplicates,TextToColumns,ConsolidateData,WhatIfAnalysisMenu,GoalSeek,DataTable,ConnectionsDialog,"
    s = s & "GetExternalDataFromText,GetExternalDataFromWeb,GetExternalDataFromAccess,SheetInsert,SheetDelete,"
    s = s & "SheetRowsInsert,SheetColumnsInsert,SheetRowsDelete,SheetColumnsDelete,RowHeight,ColumnWidth,ColumnsHide,"
    s = s & "RowsHide,ColumnsUnhide,RowsUnhide,FreezePanes,FreezeTopRow,FreezeFirstColumn,SplitWindow,ViewGridlinesToggleExcel,"
    s = s & "ZoomDialog,Zoom100,ZoomToSelection,WindowNew,WindowsArrangeAll,WindowSwitchWindowsMenuExcel,SheetProtect,"
    s = s & "ProtectWorkbook,AllowUsersToEditRanges,ReviewProtectWorkbook,ReviewShareWorkbook,ReviewNewComment,"
    s = s & "ReviewEditComment,ReviewDeleteComment,ReviewShowAllComments,TrackChanges,SpellingAndGrammar,Spelling,"
    s = s & "Thesaurus,Research,TranslationPane,FileSave,FileSaveAs,FileOpen,FileNew,FilePrint,FilePrintPreview,"
    s = s & "FilePrintQuick,FileClose,FileSendAsAttachment,FileProperties,Copy,Cut,Paste,PasteValues,PasteFormulas,"
    s = s & "PasteFormatting,PasteSpecialDialog,PasteLink,PasteAsPicture,PasteTranspose,Undo,Redo,Find,FindDialog,"
    s = s & "Replace,GoTo,GoToSpecial,SelectAll,SelectCurrentRegion,HyperlinkInsert,HyperlinkRemove,BookmarkInsert,"
    s = s & "PictureInsertFromFile,ClipArtInsert,ShapesInsertGallery,SmartArtInsert,TextBoxInsert,WordArtInsertDialog,"
    s = s & "HeaderFooterInsert,ObjectInsert,ScreenshotInsertGallery,PageSetupDialog,PageOrientationGallery,PageSizeGallery,"
    s = s & "PrintAreaMenu,PrintTitles,PageBreakInsertOrRemove,ShapeStar,ShapeHeart,ShapeSmileyFace,ShapeRectangle,"
    s = s & "ShapeRoundedRectangle,ShapeOval,ShapeRightArrow,ShapeLeftArrow,ShapeUpArrow,ShapeDownArrow,ShapeLightningBolt,"
    s = s & "ShapeSun,ShapeMoon,ShapeCloud,ShapeDiamond,ShapeTriangle,ShapeCross,HappyFace,StarRatingFull,StarRatingHalf,"
    s = s & "StarRatingEmpty,FlagRed,FlagGreen,FlagYellow,FlagBlue,FlagPurple,FlagOrange,FlagCompleted,AcceptInvitation,"
    s = s & "DeclineInvitation,TentativeAcceptInvitation,Help,Info,Tip,Lightbulb,Lock,Unlock,Key,Security,Shield,"
    s = s & "WarningIcon,ErrorIcon,HighImportance,LowImportance,ImportanceHigh,ImportanceLow,Pushpin,PinItem,CreateMailRule,"
    s = s & "NewMailMessage,MailMergeStart,AddressBook,Calendar,CalendarMonthView,CalendarWeekView,CalendarDayView,"
    s = s & "Contacts,Tasks,Notes,Journal,FolderNew,FolderOpen,OpenFolder,Folder,FolderProperties,ArchiveFolder,"
    s = s & "DatabaseQueryNew,DatabaseRelationships,DatabaseSqlServer,CreateTable,CreateQuery,CreateForm,CreateReport,"
    s = s & "GroupMacros,MacroPlay,MacroRecord,MacroSecurity,VisualBasic,ControlsGallery,DesignMode,PropertySheet,"
    s = s & "ViewCode,AddInManager,ComAddInsDialog,XmlSource,OfficeExtensionsGallery,Options,MasterViewClose,DiagramTargetInsertClassic,"
    s = s & "DiagramRadialInsertClassic,DiagramCycleInsertClassic,DiagramPyramidInsertClassic,ObjectAlignLeft,"
    s = s & "ObjectAlignCenterHorizontal,ObjectAlignRight,ObjectBringToFront,ObjectSendToBack,ObjectsGroup,ObjectRotateFree,"
    s = s & "ObjectFlipHorizontal,ObjectFlipVertical,ObjectRotateRight90,ObjectRotateLeft90,GridSettings,ViewRulerPowerPoint,"
    s = s & "ViewDraftView,ViewPrintLayoutView,ViewFullScreenView,ViewPageBreakPreviewView,ViewNormalViewExcel,"
    s = s & "ViewPageLayoutView,SlideNew,SlideShowFromBeginning,AnimationCustom,TransitionsGallery,RecordSoundInsert,"
    s = s & "MovieInsert,MediaInsert,AudioInsert,ChartRefresh,DataRefreshAll,ListMacros,TableDesign,DrawTable,"
    s = s & "TableAutoFormat,TableRowsInsertAboveExcel,TableColumnsInsertLeftExcel,ExportToPdf,FileSaveAsPdfOrXps,"
    s = s & "PublishAsPdfOrXps,FileSendMenu,ShareWorkbook,EmailAsPdf,ArrowLeft,ArrowRight,ArrowUp,ArrowDown,NextPage,"
    s = s & "PreviousPage,GoLeft,GoRight,Up,Down,MoveUp,MoveDown,CheckBox,CheckBoxInsert,OptionButton,ComboBox,"
    s = s & "ListBox,SpinButton,ScrollBar,ButtonInsert,LabelInsert,OutlineDemote,OutlinePromote,Bullets,Numbering,"
    s = s & "MultilevelListGallery,LineSpacingMenu,ParagraphMarks,"
    For Each nm In Split(s, ",")
        If Len(nm) > 0 Then
            If UI_IsOfficeIcon(CStr(nm)) Then c.Add CStr(nm)
        End If
    Next nm
    Set UI_IconCandidates = c
End Function

' True for a Digit icon that exists or an Office icon this Excel knows.
Public Function UI_IconExists(ByVal name As String) As Boolean
    If Icons_IsCustom(name) Then UI_IconExists = Icons_Exists(name) Else UI_IconExists = UI_IsOfficeIcon(name)
End Function

' 32x32 picture for the settings window (on white), or Nothing.
Public Function UI_IconPicture(ByVal name As String) As Object
    Set UI_IconPicture = Icons_FormPicture(name)
End Function

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
