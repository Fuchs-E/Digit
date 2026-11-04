Attribute VB_Name = "DigitUpdate"
'==============================================================================
' DigitUpdate - finds and installs new versions from GitHub.
'
' The "release" branch of the public repository is the update channel: its
' version.json holds the latest version and its notes. There is no prebuilt
' .xlam to download: the source on that branch is downloaded and
' build\install-update.ps1 (from that same source) builds it with Excel in the
' background, runs the self-tests, waits for this Excel to close and then
' replaces the installed add-in. If the build or a test fails, nothing is
' replaced.
'
' No UI here: DigitApp decides what to show.
'==============================================================================
Option Explicit

Public Const DIGIT_REPO As String = "Fuchs-E/Digit"
Public Const DIGIT_RELEASE_BRANCH As String = "release"
Private Const URL_VERSION As String = "https://raw.githubusercontent.com/" & DIGIT_REPO & "/" & DIGIT_RELEASE_BRANCH & "/version.json"
Private Const URL_SOURCE As String = "https://codeload.github.com/" & DIGIT_REPO & "/zip/refs/heads/" & DIGIT_RELEASE_BRANCH
Private Const USER_AGENT As String = "Digit-Excel-Addin"
Private Const REG_APP As String = "Digit"
Private Const REG_SECTION As String = "Update"

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long

' Latest published version (from version.json). Returns False with errMsg when
' it cannot be read.
Public Function Update_FetchLatest(ByRef tag As String, ByRef notes As String, ByRef pageUrl As String, _
                                   ByRef errMsg As String) As Boolean
    Dim status As Long, body As String, json As Object

    If Not HttpGetText(URL_VERSION, status, body, errMsg) Then Exit Function
    Select Case status
        Case 200
        Case 404
            errMsg = "No Digit release has been published yet."
            Exit Function
        Case 403, 429
            errMsg = "GitHub is limiting requests from this network right now. Try again in an hour."
            Exit Function
        Case Else
            errMsg = "GitHub answered with status " & status & "."
            Exit Function
    End Select

    On Error GoTo BadJson
    Set json = JsonConverter.ParseJson(body)
    tag = CStr(json("version"))
    pageUrl = "https://github.com/" & DIGIT_REPO & "/tree/" & DIGIT_RELEASE_BRANCH
    If json.Exists("notes") Then
        If Not IsNull(json("notes")) Then notes = CStr(json("notes"))
    End If
    Update_FetchLatest = (Len(tag) > 0)
    If Not Update_FetchLatest Then errMsg = "The published version.json has no version number."
    Exit Function
BadJson:
    errMsg = "Could not read GitHub's answer: " & Err.Description
End Function

' Downloads the release source and starts the background installer.
Public Function Update_Start(ByVal tag As String, ByRef errMsg As String) As Boolean
    Dim folder As String, zipPath As String, workDir As String, logPath As String, cmd As String
    Dim pid As Long

    folder = Store_Folder() & "\update"
    If Not Store_EnsureFolder(Store_Folder(), errMsg) Then Exit Function
    If Not Store_EnsureFolder(folder, errMsg) Then Exit Function
    zipPath = folder & "\digit-" & SafeName(tag) & ".zip"
    workDir = folder & "\src-" & SafeName(tag)
    logPath = folder & "\update.log"

    If Not HttpDownload(URL_SOURCE, zipPath, errMsg) Then Exit Function

    GetWindowThreadProcessId CLngPtr(Application.Hwnd), pid

    ' Unpack and hand over to the release's own installer. Paths are in single
    ' quotes (doubled inside), so the command needs no double quotes.
    cmd = "& { try { $ErrorActionPreference = 'Stop'; " & _
          "if (Test-Path -LiteralPath " & Ps(workDir) & ") { Remove-Item -LiteralPath " & Ps(workDir) & " -Recurse -Force }; " & _
          "Expand-Archive -LiteralPath " & Ps(zipPath) & " -DestinationPath " & Ps(workDir) & " -Force; " & _
          "$src = (Get-ChildItem -LiteralPath " & Ps(workDir) & " -Directory | Select-Object -First 1).FullName; " & _
          "& (Join-Path $src 'build\install-update.ps1') -Tag " & Ps(tag) & " -ExcelPid " & pid & _
          " -Dest " & Ps(ThisWorkbook.FullName) & " -LogFile " & Ps(logPath) & " } " & _
          "catch { Add-Type -AssemblyName System.Windows.Forms; " & _
          "[void][System.Windows.Forms.MessageBox]::Show('The Digit update could not start: ' + $_.Exception.Message, 'Digit') } }"

    On Error GoTo Fail
    Shell "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command """ & cmd & """", vbHide
    Update_Start = True
    Exit Function
Fail:
    errMsg = "Could not start PowerShell: " & Err.Description
End Function

' Remembered between Excel sessions (HKCU\Software\VB and VBA Program Settings\Digit\Update).
Public Function Update_GetSetting(ByVal key As String) As String
    Update_GetSetting = GetSetting(REG_APP, REG_SECTION, key, "")
End Function

Public Sub Update_SaveSetting(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting REG_APP, REG_SECTION, key, value
End Sub

'------------------------------------------------------------------------------
' HTTP. ServerXMLHTTP has timeouts; XMLHTTP is tried second because it uses the
' Windows/Edge proxy settings, which some company networks require.
'------------------------------------------------------------------------------

Private Function HttpGetText(ByVal url As String, ByRef status As Long, ByRef body As String, ByRef errMsg As String) As Boolean
    Dim http As Object
    Set http = HttpSend(url, errMsg)
    If http Is Nothing Then Exit Function
    status = http.Status
    body = http.responseText
    HttpGetText = True
End Function

Private Function HttpDownload(ByVal url As String, ByVal path As String, ByRef errMsg As String) As Boolean
    Dim http As Object, bytes() As Byte, f As Integer

    Set http = HttpSend(url, errMsg)
    If http Is Nothing Then Exit Function
    If http.Status <> 200 Then
        errMsg = "The download failed (status " & http.Status & ")."
        Exit Function
    End If
    bytes = http.responseBody
    If UBound(bytes) < 4 Then
        errMsg = "The download was empty."
        Exit Function
    End If
    If bytes(0) <> 80 Or bytes(1) <> 75 Then        ' "PK": a zip file
        errMsg = "The download is not a zip file."
        Exit Function
    End If

    On Error GoTo Fail
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , bytes
    Close #f
    HttpDownload = True
    Exit Function
Fail:
    errMsg = "Could not save the download: " & Err.Description
    On Error Resume Next
    Close #f
End Function

Private Function HttpSend(ByVal url As String, ByRef errMsg As String) As Object
    Dim progIds As Variant, i As Long, http As Object, lastErr As String

    progIds = Array("MSXML2.ServerXMLHTTP.6.0", "MSXML2.XMLHTTP.6.0")
    For i = LBound(progIds) To UBound(progIds)
        Set http = Nothing
        On Error Resume Next
        Err.Clear
        Set http = CreateObject(progIds(i))
        If Err.Number = 0 Then
            If i = 0 Then http.setTimeouts 5000, 5000, 15000, 60000
            http.Open "GET", url, False
            http.setRequestHeader "User-Agent", USER_AGENT
            http.setRequestHeader "Accept", "application/vnd.github+json"
            http.send
        End If
        If Err.Number = 0 Then
            On Error GoTo 0
            Set HttpSend = http
            Exit Function
        End If
        lastErr = Err.Description
        On Error GoTo 0
    Next i
    errMsg = "Could not reach GitHub: " & Trim$(lastErr)
End Function

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------

' PowerShell single-quoted string.
Private Function Ps(ByVal s As String) As String
    Ps = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function SafeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then out = out & ch Else out = out & "_"
    Next i
    SafeName = out
End Function
