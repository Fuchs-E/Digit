# Installing Digit

## Requirements

- Excel for Windows, Microsoft 365 (or Excel 2016 and later), 32-bit or 64-bit.
- To **build**, you also need Windows PowerShell 5.1 (it ships with Windows) and the Trust Center setting described below.

## 1. Build the add-in

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -AllowVbomAccess -RunTests
```

The script starts a hidden Excel and creates `dist\Digit.xlam`: it imports the modules, creates the Settings form, and injects the ribbon XML. Close all Excel windows before you run it.

| Option | What it does |
|---|---|
| `-AllowVbomAccess` | Turns on *Trust access to the VBA project object model* for the build only, and restores the previous setting afterwards. Without it, turn the setting on yourself: **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**. |
| `-RunTests` | Runs `DigitTests.RunAllTests` in a fresh hidden Excel and prints `PASS n` or the failures. The exit code is 1 on failure. Tests use a temporary presets file and a scratch workbook. |
| `-Install` | Copies the add-in to `%APPDATA%\Microsoft\AddIns` and ticks it in Excel's add-in list (see step 2). |
| `-ExcludeTests` | Leaves the `DigitTests` module out of the add-in. |
| `-ExportForm` | Also writes `dist\form\frmDigitSettings.frm` and `.frx`. |
| `-OutputDir <path>` | Writes the output somewhere other than `dist`. |

If you can't run the script, see [MANUAL_BUILD.md](MANUAL_BUILD.md).

## 2. Install in Excel

`-Install` does this for you. To do it by hand:

1. Copy `Digit.xlam` to a permanent folder. The usual place is `%APPDATA%\Microsoft\AddIns` (paste that into the Explorer address bar).
2. In Excel, go to **File > Options > Add-ins**. At the bottom, set **Manage** to **Excel Add-ins** and click **Go...**.
3. Click **Browse...**, select `Digit.xlam`, and click **OK**. Make sure **Digit** is ticked.

From then on, Excel loads Digit at every start, and the **Digit** group appears on the **Home** tab, right after **Number**.

## 3. Trust Center and security

Excel only runs the add-in's macros if it trusts the file. Pick one of these options:

- **Trusted location (simplest).** Add-ins in `%APPDATA%\Microsoft\AddIns` are trusted by default. If you keep the file elsewhere, add its folder under **File > Options > Trust Center > Trust Center Settings > Trusted Locations**.
- **Unblock a downloaded copy.** If the file came from email, Teams or the web, Windows marks it as coming from the internet and Excel blocks its macros. Right-click the file, choose **Properties**, tick **Unblock**, and click **OK**. Do this before you add it to Excel.
- **Sign the add-in (best for teams).** Sign the VBA project with a code-signing certificate. In the VBA editor (Alt+F11), select the Digit project and go to **Tools > Digital Signature**. Users then trust the publisher once. With *Disable all macros except digitally signed macros*, a signed add-in still runs. For personal use, you can make a test certificate with `SelfCert.exe` from the Office folder. Its signature is only trusted on your own PC.

Your IT department may enforce macro policies that override these settings.

## 4. First run

On first start, Digit creates `%APPDATA%\Digit\presets.json` with a few starter groups: Numbers, Thousands, Percent, Currency, Dates, Multiples and Inputs. Change them in **Digit > Settings**, or edit the file and click **Reload presets**. The format is described in [PRESETS.md](PRESETS.md).

## Updating

Close Excel, replace `Digit.xlam` with the new build (or run the build again with `-Install`), and start Excel. Your presets are kept because they live outside the add-in.

## Uninstalling

**File > Options > Add-ins > Manage: Excel Add-ins > Go...**, untick **Digit**, and then delete `Digit.xlam`. To remove your presets too, delete `%APPDATA%\Digit`.

## Troubleshooting

| Symptom | Fix |
|---|---|
| No Digit group on Home | Check that Digit is ticked in the Excel Add-ins dialog. If it is listed under **Disabled Items** (**File > Options > Add-ins > Manage: Disabled Items**), enable it. Check the Trust Center notes above. |
| Ribbon errors at start-up | Turn on **File > Options > Advanced > General > Show add-in user interface errors** to see the details. |
| Shortcut does nothing | Another add-in may use the same keys, or a keyboard layout may turn Ctrl+Alt into AltGr. Pick another shortcut in Settings. |
| "Presets file could not be read" | The file has a JSON error. Digit renamed it to `presets.invalid-<date>.json` and created starter presets. Fix the old file and import it from Settings. |
| Build: "Excel refused access to the VBA project" | Run with `-AllowVbomAccess`, or turn on the Trust Center setting shown in step 1. |

## Updates

Click **Digit > Check for updates**. Digit also checks once a day, about 20 seconds after Excel starts, and asks before doing anything. To turn the daily check off, set `"checkForUpdates": false` under `settings` in `presets.json`.

When you agree to an update, Digit:

1. Downloads the source of the latest version from the `release` branch of <https://github.com/Fuchs-E/Digit>.
2. Builds and tests it in the background with a separate, hidden Excel. You can keep working.
3. Waits until you close Excel, then replaces the installed `Digit.xlam`. A message confirms it.

If the build or a self-test fails, your current version is kept and a message says so. Details are in `%APPDATA%\Digit\update\update.log`. Like the first build, updating briefly turns on *Trust access to the VBA project object model* and then restores it.

