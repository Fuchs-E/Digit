Attribute VB_Name = "DigitIcons"
'==============================================================================
' DigitIcons - Digit's own button pictures ("digit:<name>" in presets.json).
'
' The build stores each PNG from src/icons on the add-in's hidden sheet
' (column E = name, column F = base64). On start-up they are written to
' %APPDATA%\Digit\icons, and GDI+ turns a PNG into a transparent picture for
' the ribbon.
'
' The settings window cannot show transparency, and assigning pictures made in
' memory to its Image controls proved unreliable, so it only ever loads plain
' BMP files with LoadPicture: Digit icons come with a white-background BMP
' (column G), and Office icons are saved once as BMP on white (Icons_FormPicture).
'==============================================================================
Option Explicit

Public Const DIGIT_ICON_PREFIX As String = "digit:"
Private Const NAME_COL As Long = 5             ' E
Private Const DATA_COL As Long = 6             ' F: PNG for the ribbon
Private Const FORM_COL As Long = 7             ' G: BMP on white for the settings window
Private Const OFFICE_CACHE As String = "office-v1"   ' change when the BMP conversion changes

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

Private Declare PtrSafe Function GetObjectW Lib "gdi32" (ByVal hObj As LongPtr, ByVal cb As Long, ByRef pv As Any) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal cb As LongPtr)
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef pictDesc As DIGIT_PICTDESC, ByRef riid As DIGIT_GUID, ByVal fOwn As Long, ByRef ppvObj As IPicture) As Long

Private mToken As LongPtr
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
    Dim ws As Worksheet, r As Long, nm As String, errMsg As String
    If StrComp(mExtractedTo, Icons_Folder(), vbTextCompare) = 0 Then Exit Sub
    mExtractedTo = Icons_Folder()
    On Error GoTo Done
    If Not Store_EnsureFolder(Store_Folder(), errMsg) Then Exit Sub
    If Not Store_EnsureFolder(Icons_Folder(), errMsg) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(1)
    r = 1
    Do While Len(CStr(ws.Cells(r, NAME_COL).Value)) > 0
        nm = CStr(ws.Cells(r, NAME_COL).Value)
        WriteIfChanged Icons_Folder() & "\" & nm & ".png", Base64Decode(CStr(ws.Cells(r, DATA_COL).Value))
        If Len(CStr(ws.Cells(r, FORM_COL).Value)) > 0 Then
            WriteIfChanged Icons_Folder() & "\" & nm & ".bmp", Base64Decode(CStr(ws.Cells(r, FORM_COL).Value))
        End If
        r = r + 1
    Loop
Done:
End Sub

' The icon as a picture, or Nothing. background: Empty keeps transparency (for
' the ribbon); a colour such as &HFFFFFF draws it on that colour (for forms,
' which do not show transparency).
Public Function Icons_Picture(ByVal nm As String, Optional ByVal background As Variant) As IPicture
    Dim bmp As LongPtr, hbm As LongPtr, argb As Long, pic As IPicture
    Dim pd As DIGIT_PICTDESC, iid As DIGIT_GUID

    If Not Icons_Exists(nm) Then Exit Function
    If IsMissing(background) Or IsEmpty(background) Then
        argb = 0
    Else
        argb = &HFF000000 Or RgbToArgb(CLng(background))
    End If
    ' No caching: the ribbon takes over a picture it is given, so handing the
    ' same picture back on the next redraw shows a broken image. A fresh one
    ' each call costs a file read of a 32x32 PNG.

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

    Set Icons_Picture = pic
End Function

' Picture for the settings window, loaded from a plain BMP on white, or Nothing.
Public Function Icons_FormPicture(ByVal nm As String) As Object
    Dim path As String, errMsg As String
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function
    On Error GoTo Done
    If Icons_IsCustom(nm) Then
        If Not Icons_Exists(nm) Then Exit Function
        path = Icons_Folder() & "\" & Mid$(nm, Len(DIGIT_ICON_PREFIX) + 1) & ".bmp"
    Else
        path = Icons_Folder() & "\" & OFFICE_CACHE & "\" & SafeFileName(nm) & ".bmp"
        If Len(Dir$(path)) = 0 Then
            If Not Store_EnsureFolder(Icons_Folder(), errMsg) Then Exit Function
            If Not Store_EnsureFolder(Icons_Folder() & "\" & OFFICE_CACHE, errMsg) Then Exit Function
            If Not SaveOfficeIconBmp(nm, path) Then Exit Function
        End If
    End If
    If Len(Dir$(path)) = 0 Then Exit Function
    Set Icons_FormPicture = LoadPicture(path)
Done:
End Function

' Saves an Office icon (32x32) as a 24-bit BMP blended onto white. GetImageMso
' returns a 32-bit DIB section with premultiplied alpha; its pixels are read
' straight from memory (GetDIBits can fail while the picture holds the bitmap).
Private Function SaveOfficeIconBmp(ByVal nm As String, ByVal path As String) As Boolean
    Dim pic As Object, h As LongPtr, bm As DIGIT_BITMAP, buf(0 To 255) As Byte, got As Long
    Dim height As Long, w As Long, hgt As Long, rowBytes As Long, src() As Byte, out() As Byte
    Dim x As Long, y As Long, srcRow As Long, si As Long, oi As Long, a As Long, c As Long, k As Long
    Dim outRow As Long, anyAlpha As Boolean, ofs As Long

    On Error GoTo Done
    Set pic = Application.CommandBars.GetImageMso(nm, 32, 32)
    If pic Is Nothing Then Exit Function
    h = CLngPtr(pic.Handle)
    If GetObjectW(h, LenB(bm), bm) = 0 Then Exit Function
    If bm.bmBitsPixel <> 32 Or bm.bmBits = 0 Then Exit Function
    w = bm.bmWidth: hgt = bm.bmHeight: rowBytes = bm.bmWidthBytes
    If w <= 0 Or hgt <= 0 Or rowBytes < w * 4 Then Exit Function

    ' Row order: DIBSECTION.dsBmih.biHeight (> 0 = bottom-up, the usual case).
    height = hgt
    got = GetObjectW(h, 256, buf(0))
#If Win64 Then
    ofs = 32 + 8
#Else
    ofs = 24 + 8
#End If
    If got >= ofs + 4 Then CopyMemory height, buf(ofs), 4

    ReDim src(0 To rowBytes * hgt - 1)
    CopyMemory src(0), ByVal bm.bmBits, rowBytes * hgt
    For k = 3 To UBound(src) Step 4
        If src(k) <> 0 Then anyAlpha = True: Exit For
    Next k

    outRow = ((w * 3 + 3) \ 4) * 4
    ReDim out(0 To 54 + outRow * hgt - 1)
    out(0) = 66: out(1) = 77                                   ' "BM"
    PutLong out, 2, 54 + outRow * hgt
    PutLong out, 10, 54
    PutLong out, 14, 40
    PutLong out, 18, w
    PutLong out, 22, hgt                                       ' bottom-up
    out(26) = 1: out(28) = 24
    PutLong out, 34, outRow * hgt

    For y = 0 To hgt - 1                                       ' y = 0 is the bottom row in the file
        If height > 0 Then srcRow = y Else srcRow = hgt - 1 - y
        For x = 0 To w - 1
            si = srcRow * rowBytes + x * 4
            oi = 54 + y * outRow + x * 3
            If anyAlpha Then a = src(si + 3) Else a = 255
            For k = 0 To 2                                     ' B, G, R
                c = CLng(src(si + k)) + (255 - a)
                If c > 255 Then c = 255
                out(oi + k) = c
            Next k
        Next x
    Next y

    WriteIfChanged path, out
    SaveOfficeIconBmp = (Len(Dir$(path)) > 0)
Done:
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal at As Long, ByVal v As Long)
    CopyMemory b(at), v, 4
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_-]" Then out = out & ch Else out = out & "~"
    Next i
    SafeFileName = out
End Function

Private Sub WriteIfChanged(ByVal path As String, ByRef bytes() As Byte)
    Dim f As Integer
    If FileSize(path) = UBound(bytes) - LBound(bytes) + 1 Then Exit Sub
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , bytes
    Close #f
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
