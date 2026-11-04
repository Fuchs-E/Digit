Attribute VB_Name = "DigitIcons"
'==============================================================================
' DigitIcons - Digit's own button pictures ("digit:<name>" in presets.json).
'
' The build stores each PNG from src/icons on the add-in's hidden sheet
' (column E = name, column F = base64). On start-up they are written to
' %APPDATA%\Digit\icons, and GDI+ turns a PNG into a picture the ribbon and the
' settings window can show. Built-in Office icons (imageMso names) are handled
' in DigitUI; UI_IconExists covers both kinds.
'==============================================================================
Option Explicit

Public Const DIGIT_ICON_PREFIX As String = "digit:"
Private Const NAME_COL As Long = 5             ' E
Private Const DATA_COL As Long = 6             ' F

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type DIGIT_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type DIGIT_PICTDESC
    cbSizeofstruct As Long
    picType As Long
    hImage As LongPtr
    xExt As LongPtr
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Type DIGIT_BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
End Type

Private Type DIGIT_BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Declare PtrSafe Function GetObjectW Lib "gdi32" (ByVal hObj As LongPtr, ByVal cb As Long, ByRef pv As Any) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hbm As LongPtr, ByVal uStart As Long, ByVal cLines As Long, ByRef lpvBits As Any, ByRef lpbmi As DIGIT_BITMAPINFOHEADER, ByVal usage As Long) As Long
Private Declare PtrSafe Function SetDIBits Lib "gdi32" (ByVal hdc As LongPtr, ByVal hbm As LongPtr, ByVal uStart As Long, ByVal cLines As Long, ByRef lpvBits As Any, ByRef lpbmi As DIGIT_BITMAPINFOHEADER, ByVal usage As Long) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal cx As Long, ByVal cy As Long) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef pictDesc As DIGIT_PICTDESC, ByRef riid As DIGIT_GUID, ByVal fOwn As Long, ByRef ppvObj As IPicture) As Long

Private mToken As LongPtr
Private mCache As Scripting.Dictionary        ' "name|background" -> IPicture
Private mExtractedTo As String               ' folder the icons were last written to

' True for "digit:..." names.
Public Function Icons_IsCustom(ByVal nm As String) As Boolean
    Icons_IsCustom = (LCase$(Left$(Trim$(nm), Len(DIGIT_ICON_PREFIX))) = DIGIT_ICON_PREFIX)
End Function

Public Function Icons_Folder() As String
    Icons_Folder = Store_Folder() & "\icons"
End Function

Private Function IconFile(ByVal nm As String) As String
    IconFile = Icons_Folder() & "\" & Mid$(Trim$(nm), Len(DIGIT_ICON_PREFIX) + 1) & ".png"
End Function

' True when the custom icon exists (after extraction).
Public Function Icons_Exists(ByVal nm As String) As Boolean
    If Not Icons_IsCustom(nm) Then Exit Function
    Icons_Extract
    On Error Resume Next
    Icons_Exists = (Len(Dir$(IconFile(nm))) > 0)
End Function

' Names of every custom icon packed in the add-in ("digit:..."), in sheet order.
Public Function Icons_List() As Collection
    Dim ws As Worksheet, r As Long, c As New Collection
    Set ws = ThisWorkbook.Worksheets(1)
    r = 1
    Do While Len(CStr(ws.Cells(r, NAME_COL).Value)) > 0
        c.Add DIGIT_ICON_PREFIX & CStr(ws.Cells(r, NAME_COL).Value)
        r = r + 1
    Loop
    Set Icons_List = c
End Function

' Writes the packed icons to the icons folder (once per session; only files that
' are missing or differ in size are rewritten).
Public Sub Icons_Extract()
    Dim ws As Worksheet, r As Long, nm As String, bytes() As Byte, path As String, f As Integer, errMsg As String
    If StrComp(mExtractedTo, Icons_Folder(), vbTextCompare) = 0 Then Exit Sub
    mExtractedTo = Icons_Folder()
    Set mCache = Nothing
    On Error GoTo Done
    If Not Store_EnsureFolder(Store_Folder(), errMsg) Then Exit Sub
    If Not Store_EnsureFolder(Icons_Folder(), errMsg) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(1)
    r = 1
    Do While Len(CStr(ws.Cells(r, NAME_COL).Value)) > 0
        nm = CStr(ws.Cells(r, NAME_COL).Value)
        bytes = Base64Decode(CStr(ws.Cells(r, DATA_COL).Value))
        path = Icons_Folder() & "\" & nm & ".png"
        If FileSize(path) <> UBound(bytes) + 1 Then
            If Len(Dir$(path)) > 0 Then Kill path
            f = FreeFile
            Open path For Binary Access Write As #f
            Put #f, , bytes
            Close #f
        End If
        r = r + 1
    Loop
Done:
End Sub

' The icon as a picture, or Nothing. background: Empty keeps transparency (for
' the ribbon); a colour such as &HFFFFFF draws it on that colour (for forms,
' which do not show transparency).
Public Function Icons_Picture(ByVal nm As String, Optional ByVal background As Variant) As IPicture
    Dim key As String, bmp As LongPtr, hbm As LongPtr, argb As Long, pic As IPicture
    Dim pd As DIGIT_PICTDESC, iid As DIGIT_GUID

    If Not Icons_Exists(nm) Then Exit Function
    If IsMissing(background) Or IsEmpty(background) Then
        argb = 0
    Else
        argb = &HFF000000 Or RgbToArgb(CLng(background))
    End If
    key = LCase$(nm) & "|" & argb
    If mCache Is Nothing Then Set mCache = New Scripting.Dictionary
    If mCache.Exists(key) Then
        Set Icons_Picture = mCache(key)
        Exit Function
    End If

    If Not StartGdiPlus() Then Exit Function
    If GdipCreateBitmapFromFile(StrPtr(IconFile(nm)), bmp) <> 0 Then Exit Function
    If GdipCreateHBITMAPFromBitmap(bmp, hbm, argb) <> 0 Then
        GdipDisposeImage bmp
        Exit Function
    End If
    GdipDisposeImage bmp

    pd.cbSizeofstruct = LenB(pd)
    pd.picType = 1                              ' PICTYPE_BITMAP
    pd.hImage = hbm
    ' IID_IPicture {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    iid.Data1 = &H7BF80980: iid.Data2 = &HBF32: iid.Data3 = &H101A
    iid.Data4(0) = &H8B: iid.Data4(1) = &HBB: iid.Data4(2) = &H0: iid.Data4(3) = &HAA
    iid.Data4(4) = &H0: iid.Data4(5) = &H30: iid.Data4(6) = &HC: iid.Data4(7) = &HAB
    If OleCreatePictureIndirect(pd, iid, 1, pic) <> 0 Then Exit Function

    Set mCache(key) = pic
    Set Icons_Picture = pic
End Function

' A copy of pic (a 32-bit picture with transparency, such as GetImageMso
' returns) blended onto a solid colour. UserForm Image controls ignore
' transparency and show those areas black. Returns pic itself when it has no
' transparency or cannot be read.
Public Function Icons_Flatten(ByVal pic As Object, ByVal bgColor As Long) As Object
    Dim bm As DIGIT_BITMAP, bi As DIGIT_BITMAPINFOHEADER, hdc As LongPtr, hbm As LongPtr, src As LongPtr
    Dim px() As Long, i As Long, n As Long, v As Long, a As Long, r As Long, g As Long, b As Long
    Dim bgR As Long, bgG As Long, bgB As Long, anyAlpha As Boolean
    Dim pd As DIGIT_PICTDESC, iid As DIGIT_GUID, out As IPicture

    Set Icons_Flatten = pic
    If pic Is Nothing Then Exit Function
    On Error GoTo Done
    src = CLngPtr(pic.Handle)
    If GetObjectW(src, LenB(bm), bm) = 0 Then Exit Function
    If bm.bmBitsPixel <> 32 Or bm.bmWidth <= 0 Or bm.bmHeight <= 0 Then Exit Function

    bi.biSize = LenB(bi)
    bi.biWidth = bm.bmWidth
    bi.biHeight = -bm.bmHeight                  ' top-down
    bi.biPlanes = 1
    bi.biBitCount = 32
    n = bm.bmWidth * bm.bmHeight
    ReDim px(0 To n - 1)
    hdc = GetDC(0)
    If GetDIBits(hdc, src, 0, bm.bmHeight, px(0), bi, 0) = 0 Then ReleaseDC 0, hdc: Exit Function

    For i = 0 To n - 1
        If (px(i) And &HFF000000) <> 0 Then anyAlpha = True: Exit For
    Next i
    If Not anyAlpha Then ReleaseDC 0, hdc: Exit Function

    bgR = bgColor And &HFF&: bgG = (bgColor \ &H100&) And &HFF&: bgB = (bgColor \ &H10000) And &HFF&
    For i = 0 To n - 1
        v = px(i)
        a = (v And &H7F000000) \ &H1000000
        If v < 0 Then a = a Or &H80&
        b = v And &HFF&: g = (v \ &H100&) And &HFF&: r = (v And &HFF0000) \ &H10000
        ' Colours are premultiplied by alpha: add the background where it shows.
        r = r + ((255 - a) * bgR) \ 255: If r > 255 Then r = 255
        g = g + ((255 - a) * bgG) \ 255: If g > 255 Then g = 255
        b = b + ((255 - a) * bgB) \ 255: If b > 255 Then b = 255
        px(i) = b Or (g * &H100&) Or (r * &H10000)
    Next i

    hbm = CreateCompatibleBitmap(hdc, bm.bmWidth, bm.bmHeight)
    If hbm = 0 Then ReleaseDC 0, hdc: Exit Function
    SetDIBits hdc, hbm, 0, bm.bmHeight, px(0), bi, 0
    ReleaseDC 0, hdc

    pd.cbSizeofstruct = LenB(pd)
    pd.picType = 1
    pd.hImage = hbm
    iid.Data1 = &H7BF80980: iid.Data2 = &HBF32: iid.Data3 = &H101A
    iid.Data4(0) = &H8B: iid.Data4(1) = &HBB: iid.Data4(2) = &H0: iid.Data4(3) = &HAA
    iid.Data4(4) = &H0: iid.Data4(5) = &H30: iid.Data4(6) = &HC: iid.Data4(7) = &HAB
    If OleCreatePictureIndirect(pd, iid, 1, out) <> 0 Then Exit Function
    Set Icons_Flatten = out
Done:
End Function

Public Sub Icons_ClearCache()
    Set mCache = Nothing
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------

Private Function StartGdiPlus() As Boolean
    Dim si As GdiplusStartupInput
    If mToken <> 0 Then StartGdiPlus = True: Exit Function
    si.GdiplusVersion = 1
    StartGdiPlus = (GdiplusStartup(mToken, si, 0) = 0)
End Function

' VBA colour (&HBBGGRR) to GDI+ (&HRRGGBB, alpha added by the caller).
Private Function RgbToArgb(ByVal c As Long) As Long
    RgbToArgb = ((c And &HFF&) * &H10000) Or (c And &HFF00&) Or ((c And &HFF0000) \ &H10000)
End Function

Private Function Base64Decode(ByVal s As String) As Byte()
    Dim doc As Object, node As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = doc.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = s
    Base64Decode = node.nodeTypedValue
End Function

Private Function FileSize(ByVal path As String) As Long
    FileSize = -1
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then FileSize = FileLen(path)
End Function
