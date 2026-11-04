<#
.SYNOPSIS
    Installs a Digit update. Started by Digit's "Check for updates" button.

.DESCRIPTION
    Runs from the unpacked source of the new release, hidden:
    1. Checks that the source's DIGIT_VERSION matches the release tag.
    2. Builds Digit.xlam with build.ps1 into a staging folder and runs the
       self-tests (a separate hidden Excel; the user's Excel is not touched).
    3. Waits until the user's Excel (ExcelPid) has closed.
    4. Copies the new add-in over the installed one (Dest).
    The result is shown in a message box and written to LogFile. If anything
    fails before step 4, the installed add-in is left as it was.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Tag,
    [Parameter(Mandatory)] [int]$ExcelPid,
    [Parameter(Mandatory)] [string]$Dest,
    [Parameter(Mandatory)] [string]$LogFile
)

$ErrorActionPreference = 'Stop'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Add-Type -AssemblyName System.Windows.Forms

function Log([string]$text) {
    $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $text
    Add-Content -LiteralPath $LogFile -Value $line
}

function Show([string]$text, [string]$icon = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($text, 'Digit', 'OK', $icon)
}

try {
    Log "---- Update to $Tag (Excel process $ExcelPid, add-in $Dest)"

    # 1. The release must contain the version it claims, or Digit would offer
    #    the same update again after installing it.
    $appSource = Get-Content -LiteralPath (Join-Path $Root 'src\vba\DigitApp.bas') -Raw
    if ($appSource -notmatch 'DIGIT_VERSION As String = "([^"]+)"') { throw 'DIGIT_VERSION not found in the release source.' }
    $version = $Matches[1]
    if ($Tag.TrimStart('v', 'V') -ne $version) {
        throw "Release $Tag contains version $version. The release is mislabelled, so it was not installed."
    }

    # 2. Build and test.
    $staging = Join-Path (Split-Path $LogFile) "build-$version"
    Log "Building into $staging"
    $output = & (Join-Path $Root 'build\build.ps1') -OutputDir $staging -AllowVbomAccess -RunTests *>&1 | Out-String
    $buildExit = $LASTEXITCODE
    Log $output
    $built = Join-Path $staging 'Digit.xlam'
    if ($buildExit -ne 0 -or -not (Test-Path -LiteralPath $built)) {
        throw "The build or its self-tests failed, so your current Digit was kept. Details: $LogFile"
    }
    Log 'Build and tests passed.'

    # 3. Wait for the user's Excel to close (the add-in file is in use until then).
    $proc = Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue
    if ($null -ne $proc) {
        Log 'Waiting for Excel to close.'
        $proc.WaitForExit()
    }

    # 4. Replace the add-in. Excel can take a moment to release the file.
    $copied = $false
    for ($i = 0; $i -lt 30 -and -not $copied; $i++) {
        try { Copy-Item -LiteralPath $built -Destination $Dest -Force; $copied = $true }
        catch { Start-Sleep -Seconds 2 }
    }
    if (-not $copied) { throw "Could not replace $Dest. Is another Excel window still open? Close it and click Check for updates again." }

    Log "Installed $version."
    Show "Digit $version is installed. It will be ready the next time you start Excel."
}
catch {
    $msg = $_.Exception.Message
    try { Log "FAILED: $msg" } catch { }
    # Let "Check for updates" offer this version again.
    Remove-ItemProperty -Path 'HKCU:\Software\VB and VBA Program Settings\Digit\Update' -Name Pending -ErrorAction SilentlyContinue
    Show "The Digit update did not complete.`n`n$msg" 'Warning'
    exit 1
}
