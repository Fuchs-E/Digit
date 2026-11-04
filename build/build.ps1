<#
.SYNOPSIS
    Builds Digit.xlam from the plain-text sources using Excel COM automation.

.DESCRIPTION
    1. Starts a hidden Excel and creates a one-sheet workbook.
    2. Adds the "Microsoft Scripting Runtime" reference (needed by VBA-JSON).
    3. Imports the standard modules (src\vba\*.bas and lib\JsonConverter.bas).
    4. Pastes ThisWorkbook.code.vba into the workbook's ThisWorkbook module.
    5. Creates the frmDigitSettings UserForm with build\DigitFormBuilder.bas.
    6. Saves the workbook as an add-in (.xlam).
    7. Injects src\ribbon\customUI14.xml into the .xlam package.
    Optionally runs the self-tests and installs the add-in for the current user.

    Requires Excel for Windows and "Trust access to the VBA project object model"
    (File > Options > Trust Center > Trust Center Settings > Macro Settings).
    Use -AllowVbomAccess to switch it on just for the build.

.PARAMETER OutputDir
    Where Digit.xlam is written. Default: <repo>\dist

.PARAMETER RunTests
    Opens the built add-in in a fresh hidden Excel and runs DigitTests.RunAllTests.
    Exit code 1 when a test fails.

.PARAMETER Install
    Copies Digit.xlam to %APPDATA%\Microsoft\AddIns and ticks it in Excel's
    add-in list so it loads at every start. Close Excel first.

.PARAMETER ExcludeTests
    Leaves the DigitTests module out of the add-in (it is small and harmless,
    but you may prefer to ship without it). Cannot be combined with -RunTests.

.PARAMETER ExportForm
    Also writes the generated UserForm to dist\form\frmDigitSettings.frm/.frx,
    for anyone who prefers to import the form by hand.

.PARAMETER AllowVbomAccess
    Temporarily turns on "Trust access to the VBA project object model" for the
    current user and restores the previous setting afterwards.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build.ps1 -RunTests -Install
#>
[CmdletBinding()]
param(
    [string]$OutputDir,
    [switch]$RunTests,
    [switch]$Install,
    [switch]$ExcludeTests,
    [switch]$ExportForm,
    [switch]$AllowVbomAccess
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not $OutputDir) { $OutputDir = Join-Path $Root 'dist' }
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$XlamPath = Join-Path $OutputDir 'Digit.xlam'
$TempDir = Join-Path ([IO.Path]::GetTempPath()) ("digit-build-" + [Guid]::NewGuid().ToString('N'))

if ($RunTests -and $ExcludeTests) { throw '-RunTests needs the DigitTests module; drop -ExcludeTests.' }

# Excel constants
$xlOpenXMLAddIn = 55
$vbextCtStdModule = 1
$vbextCtMSForm = 3
$vbextCtDocument = 100
$msoAutomationSecurityForceDisable = 3

function Write-Step([string]$text) { Write-Host "==> $text" -ForegroundColor Cyan }

# VBA imports text in the ANSI code page and expects CRLF line endings. Git on
# Windows may check files out with LF, so normalise into a temporary copy.
function Get-CrlfCopy([string]$path) {
    $text = [IO.File]::ReadAllText($path)
    $text = $text -replace "`r?`n", "`r`n"
    foreach ($ch in $text.ToCharArray()) {
        if ([int]$ch -gt 127) { throw "Non-ASCII character in $path. Keep VBA sources ASCII-only." }
    }
    $dest = Join-Path $TempDir (Split-Path $path -Leaf)
    [IO.File]::WriteAllText($dest, $text, [Text.Encoding]::ASCII)
    return $dest
}

function Release-Com($obj) {
    if ($null -ne $obj) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch { }
    }
}

#------------------------------------------------------------------------------
# "Trust access to the VBA project object model"
#------------------------------------------------------------------------------
$SecurityKey = $null
$PreviousVbom = $null
$VbomChanged = $false

function Get-ExcelSecurityKey {
    foreach ($ver in '16.0', '15.0') {
        $k = "HKCU:\Software\Microsoft\Office\$ver\Excel"
        if (Test-Path $k) { return "$k\Security" }
    }
    return 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
}

function Enable-VbomAccess {
    $script:SecurityKey = Get-ExcelSecurityKey
    if (-not (Test-Path $script:SecurityKey)) { New-Item -Path $script:SecurityKey -Force | Out-Null }
    $current = Get-ItemProperty -Path $script:SecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -ne $current) { $script:PreviousVbom = $current.AccessVBOM }
    if ($script:PreviousVbom -ne 1) {
        Set-ItemProperty -Path $script:SecurityKey -Name AccessVBOM -Value 1 -Type DWord
        $script:VbomChanged = $true
        Write-Host '    Trust access to the VBA project object model: enabled for this build.'
    }
}

function Restore-VbomAccess {
    if (-not $script:VbomChanged) { return }
    if ($null -eq $script:PreviousVbom) {
        Remove-ItemProperty -Path $script:SecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $script:SecurityKey -Name AccessVBOM -Value $script:PreviousVbom -Type DWord
    }
    Write-Host '    Trust access to the VBA project object model: restored.'
}

#------------------------------------------------------------------------------
# customUI injection (an .xlam is a zip package)
#------------------------------------------------------------------------------
# Ribbon images: every src\ribbon\images\<name>.png becomes customUI/images/<name>.png,
# referenced from customUI14.xml as image="<name>".
function Add-CustomUI([string]$packagePath, [string]$customUiPath, [string]$imagesDir) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $partName = 'customUI/customUI14.xml'
    $relType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
    $relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $utf8 = New-Object Text.UTF8Encoding($false)

    # Validate the XML before touching the package.
    $uiXml = New-Object Xml.XmlDocument
    $uiXml.Load($customUiPath)

    $zip = [IO.Compression.ZipFile]::Open($packagePath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $old = $zip.GetEntry($partName)
        if ($null -ne $old) { $old.Delete() }
        $entry = $zip.CreateEntry($partName)
        $bytes = [IO.File]::ReadAllBytes($customUiPath)
        $stream = $entry.Open()
        try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }

        # Custom images and the customUI part's own relationships file.
        $pngs = @()
        if ($imagesDir -and (Test-Path $imagesDir)) { $pngs = @(Get-ChildItem $imagesDir -Filter '*.png') }
        if ($pngs.Count -gt 0) {
            $imgRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            $sb = New-Object Text.StringBuilder
            [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            [void]$sb.Append("<Relationships xmlns=`"$relNs`">")
            foreach ($png in $pngs) {
                $imgPart = "customUI/images/$($png.Name)"
                $oldImg = $zip.GetEntry($imgPart)
                if ($null -ne $oldImg) { $oldImg.Delete() }
                $imgEntry = $zip.CreateEntry($imgPart)
                $imgBytes = [IO.File]::ReadAllBytes($png.FullName)
                $s = $imgEntry.Open()
                try { $s.Write($imgBytes, 0, $imgBytes.Length) } finally { $s.Dispose() }
                [void]$sb.Append("<Relationship Id=`"$($png.BaseName)`" Type=`"$imgRelType`" Target=`"images/$($png.Name)`"/>")
            }
            [void]$sb.Append('</Relationships>')
            $uiRelsPart = 'customUI/_rels/customUI14.xml.rels'
            $oldRels = $zip.GetEntry($uiRelsPart)
            if ($null -ne $oldRels) { $oldRels.Delete() }
            $relBytes = $utf8.GetBytes($sb.ToString())
            $s = $zip.CreateEntry($uiRelsPart).Open()
            try { $s.Write($relBytes, 0, $relBytes.Length) } finally { $s.Dispose() }
        }

        # Package relationship pointing at the ribbon part.
        $relsEntry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object IO.StreamReader($relsEntry.Open())
        try { $relsText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $rels = New-Object Xml.XmlDocument
        $rels.LoadXml($relsText)
        $found = $false
        foreach ($node in @($rels.DocumentElement.ChildNodes)) {
            if ($node.NodeType -eq [Xml.XmlNodeType]::Element -and $node.GetAttribute('Type') -eq $relType) {
                $node.SetAttribute('Target', $partName)
                $found = $true
            }
        }
        if (-not $found) {
            $rel = $rels.CreateElement('Relationship', $relNs)
            $rel.SetAttribute('Id', 'rIdDigitCustomUI')
            $rel.SetAttribute('Type', $relType)
            $rel.SetAttribute('Target', $partName)
            [void]$rels.DocumentElement.AppendChild($rel)
        }
        $relsEntry.Delete()
        $newRels = $zip.CreateEntry('_rels/.rels')
        $writer = New-Object IO.StreamWriter($newRels.Open(), $utf8)
        try { $rels.Save($writer) } finally { $writer.Dispose() }

        # Make sure .xml parts have a content type (Excel always writes it, but be safe).
        $ctEntry = $zip.GetEntry('[Content_Types].xml')
        $reader = New-Object IO.StreamReader($ctEntry.Open())
        try { $ctText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $ct = New-Object Xml.XmlDocument
        $ct.LoadXml($ctText)
        $ctChanged = $false
        foreach ($pair in @(@('xml', 'application/xml'), @('png', 'image/png'))) {
            $has = $false
            foreach ($node in @($ct.DocumentElement.ChildNodes)) {
                if ($node.NodeType -eq [Xml.XmlNodeType]::Element -and $node.LocalName -eq 'Default' -and
                    $node.GetAttribute('Extension') -eq $pair[0]) { $has = $true }
            }
            if (-not $has) {
                $def = $ct.CreateElement('Default', $ct.DocumentElement.NamespaceURI)
                $def.SetAttribute('Extension', $pair[0])
                $def.SetAttribute('ContentType', $pair[1])
                [void]$ct.DocumentElement.PrependChild($def)
                $ctChanged = $true
            }
        }
        if ($ctChanged) {
            $ctEntry.Delete()
            $newCt = $zip.CreateEntry('[Content_Types].xml')
            $writer = New-Object IO.StreamWriter($newCt.Open(), $utf8)
            try { $ct.Save($writer) } finally { $writer.Dispose() }
        }
    } finally {
        $zip.Dispose()
    }
}

#------------------------------------------------------------------------------
# Source check: VBA only allows declarations (Const, Dim, Declare, Type, Enum)
# before the first procedure. Excel reports a misplaced one as a compile error
# only when the module is first used, which the self-tests may never do, so
# catch it here with the file and line.
#------------------------------------------------------------------------------
function Test-VbaLayout([string[]]$paths) {
    $problems = @()
    foreach ($path in $paths) {
        $inProcs = $false
        $n = 0
        foreach ($line in [IO.File]::ReadAllLines($path)) {
            $n++
            if ($line -match '^\s*((Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property)\s') { $inProcs = $true; continue }
            if ($inProcs -and $line -match '^((Public|Private|Global)\s+)?(Const|Dim|Declare|Type|Enum)\s') {
                $problems += "$(Split-Path $path -Leaf) line ${n}: $($line.Trim())"
            }
        }
    }
    if ($problems.Count -gt 0) {
        throw ("Declarations must come before the first Sub or Function (VBA compile error):`n    " + ($problems -join "`n    "))
    }
}

#------------------------------------------------------------------------------
# Build
#------------------------------------------------------------------------------
$excel = $null
$wb = $null
$exitCode = 0

New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

try {
    Write-Step 'Checking sources'
    Test-VbaLayout (@(Get-ChildItem (Join-Path $Root 'src\vba') -Filter '*.bas' | ForEach-Object FullName) +
                    @(Get-ChildItem (Join-Path $Root 'src') -Recurse -Filter '*.code.vba' | ForEach-Object FullName))

    if ($AllowVbomAccess) { Enable-VbomAccess }

    if (Test-Path $XlamPath) {
        try { Remove-Item $XlamPath -Force }
        catch { throw "Cannot overwrite $XlamPath. Is it open in Excel? Close Excel and try again." }
    }

    Write-Step 'Starting Excel'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    Write-Step 'Creating workbook'
    $wb = $excel.Workbooks.Add(-4167)          # xlWBATWorksheet: one sheet
    $wb.Worksheets.Item(1).Name = 'Digit'

    try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count }
    catch {
        throw ("Excel refused access to the VBA project. Turn on File > Options > Trust Center > " +
               "Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', " +
               "or run this script with -AllowVbomAccess.")
    }
    $vbp.Name = 'Digit'

    Write-Step 'Adding reference: Microsoft Scripting Runtime'
    $hasScripting = $false
    foreach ($ref in $vbp.References) { if ($ref.Name -eq 'Scripting') { $hasScripting = $true } }
    if (-not $hasScripting) {
        # Try the registered type library first (exact version, then latest), and fall back to
        # loading scrrun.dll directly. Some PCs report "Object library not registered" for the
        # GUID lookup. In 32-bit Excel, System32 is redirected to SysWOW64 automatically.
        $guid = '{420B2830-E718-11CF-893D-00C04FD7AC20}'
        $dll  = Join-Path $env:windir 'System32\scrrun.dll'
        $attempts = @(
            @{ Name = 'GUID 1.0';    Run = { [void]$vbp.References.AddFromGuid($guid, 1, 0) } },
            @{ Name = 'GUID latest'; Run = { [void]$vbp.References.AddFromGuid($guid, 0, 0) } },
            @{ Name = $dll;          Run = { [void]$vbp.References.AddFromFile($dll) } }
        )
        $added = $false; $errors = @()
        foreach ($a in $attempts) {
            try { & $a.Run; $added = $true; Write-Host "    added via $($a.Name)"; break }
            catch { $errors += "$($a.Name): $($_.Exception.Message)" }
        }
        if (-not $added) {
            throw ("Could not add the Microsoft Scripting Runtime reference.`n    " + ($errors -join "`n    "))
        }
    }

    Write-Step 'Importing modules'
    $modules = @(Get-Item (Join-Path $Root 'lib\JsonConverter.bas')) +
               @(Get-ChildItem (Join-Path $Root 'src\vba') -Filter '*.bas' | Sort-Object Name)
    foreach ($m in $modules) {
        if ($ExcludeTests -and $m.BaseName -eq 'DigitTests') { continue }
        Write-Host "    $($m.Name)"
        [void]$vbp.VBComponents.Import((Get-CrlfCopy $m.FullName))
    }

    Write-Step 'Adding ThisWorkbook code'
    $docComp = $null
    $codeName = $wb.CodeName                   # localised, e.g. "DieseArbeitsmappe"
    foreach ($c in $vbp.VBComponents) {
        if ($c.Type -eq $vbextCtDocument -and $c.Name -eq $codeName) { $docComp = $c }
    }
    if ($null -eq $docComp) {
        foreach ($c in $vbp.VBComponents) {
            if ($c.Type -eq $vbextCtDocument) {
                try { $null = $c.Properties.Item('IsAddin'); $docComp = $c; break } catch { }
            }
        }
    }
    if ($null -eq $docComp) { throw 'Could not find the ThisWorkbook module.' }
    $cm = $docComp.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString([IO.File]::ReadAllText((Get-CrlfCopy (Join-Path $Root 'src\vba\ThisWorkbook.code.vba'))))

    Write-Step 'Building the settings form'
    [void]$vbp.VBComponents.Import((Get-CrlfCopy (Join-Path $Root 'build\DigitFormBuilder.bas')))
    $formCode = Get-CrlfCopy (Join-Path $Root 'src\forms\frmDigitSettings.code.vba')
    $excel.Run("'" + $wb.Name + "'!DigitFormBuilder.BuildSettingsForm", $formCode)
    $vbp.VBComponents.Remove($vbp.VBComponents.Item('DigitFormBuilder'))

    if ($ExportForm) {
        $formDir = Join-Path $OutputDir 'form'
        New-Item -ItemType Directory -Path $formDir -Force | Out-Null
        $vbp.VBComponents.Item('frmDigitSettings').Export((Join-Path $formDir 'frmDigitSettings.frm'))
        Write-Host "    Exported to $formDir"
    }

    Write-Step 'Packing Digit icons'
    # Each src\icons\<name>.png goes on the hidden sheet: column E = name,
    # column F = PNG (base64), column G = form\<name>.bmp (base64).
    # DigitIcons writes them to %APPDATA%\Digit\icons.
    $iconRow = 0
    $iconSheet = $wb.Worksheets.Item(1)
    foreach ($png in @(Get-ChildItem (Join-Path $Root 'src\icons') -Filter '*.png' | Sort-Object Name)) {
        $iconRow++
        $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($png.FullName))
        if ($b64.Length -gt 32000) { throw "$($png.Name) is too large to pack (keep icons small, 32x32)." }
        $iconSheet.Range("E$iconRow").Value2 = $png.BaseName
        $iconSheet.Range("F$iconRow").Value2 = $b64
        # Column G: the same icon on white as BMP, for the settings window.
        $bmp = Join-Path $png.DirectoryName ("form\" + $png.BaseName + '.bmp')
        if (-not (Test-Path $bmp)) { throw "Missing $bmp. Run src\icons\make_icons.py." }
        $iconSheet.Range("G$iconRow").Value2 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($bmp))
    }
    Write-Host "    $iconRow icons"

    Write-Step 'Saving add-in'
    # DocumentProperties has no type info, so PowerShell's normal member access returns $null.
    # Go through IDispatch by hand. These only label the add-in in Excel's Add-ins list, so a
    # failure here is reported but does not stop the build.
    try {
        $props = $wb.BuiltinDocumentProperties
        foreach ($kv in @(@('Title', 'Digit'),
                          @('Comments', 'One-click number format presets. Home tab > Digit group.'))) {
            $prop = [System.__ComObject].InvokeMember('Item', [System.Reflection.BindingFlags]::GetProperty, $null, $props, @($kv[0]))
            [void][System.__ComObject].InvokeMember('Value', [System.Reflection.BindingFlags]::SetProperty, $null, $prop, @($kv[1]))
        }
    } catch { Write-Host "    (could not set the add-in title: $($_.Exception.Message))" -ForegroundColor Yellow }
    $wb.SaveAs($XlamPath, $xlOpenXMLAddIn)
    $wb.Close($false)
    Release-Com $wb; $wb = $null
    $excel.Quit()
    Release-Com $excel; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    Write-Step 'Injecting ribbon (customUI14.xml)'
    Add-CustomUI $XlamPath (Join-Path $Root 'src\ribbon\customUI14.xml') (Join-Path $Root 'src\ribbon\images')

    Write-Host ''
    Write-Host "Built $XlamPath" -ForegroundColor Green

    #--------------------------------------------------------------------------
    if ($RunTests) {
        Write-Step 'Running self-tests'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false           # skip Workbook_Open: tests use their own presets file
        $wb = $excel.Workbooks.Open($XlamPath)
        $result = [string]$excel.Run("'Digit.xlam'!DigitTests.RunAllTests", $TempDir)
        Write-Host $result
        if ($result -notmatch '^PASS') { $exitCode = 1 }
        $wb.Close($false)
        Release-Com $wb; $wb = $null
        $excel.Quit()
        Release-Com $excel; $excel = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    #--------------------------------------------------------------------------
    if ($Install) {
        Write-Step 'Installing for the current user'
        $addinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
        New-Item -ItemType Directory -Path $addinsDir -Force | Out-Null
        $dest = Join-Path $addinsDir 'Digit.xlam'
        try { Copy-Item $XlamPath $dest -Force }
        catch { throw "Cannot replace $dest. Close Excel and run the build again." }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $tmpBook = $excel.Workbooks.Add()       # AddIns.Add needs an open workbook
        $addin = $excel.AddIns.Add($dest, $false)
        $addin.Installed = $true
        $tmpBook.Close($false)
        Release-Com $addin; Release-Com $tmpBook
        $excel.Quit()
        Release-Com $excel; $excel = $null
        Write-Host "Installed $dest. Start Excel and look for the Digit group on the Home tab." -ForegroundColor Green
    }
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } ; Release-Com $wb }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } ; Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Restore-VbomAccess
    Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
}

exit $exitCode
