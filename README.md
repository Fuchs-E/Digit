# Digit

Digit is an Excel add-in for Windows desktop (Microsoft 365). It adds a **Digit** tab to the ribbon with one-click buttons that apply your own formatting presets to the selected cells.

Each quick button holds a **preset group** of up to three presets. Click it again and it moves to the next preset: 1 → 2 → 3 → 1. A preset is a number format plus an optional font colour and fill colour. Nothing is built in. You decide what the groups are and which ones go on the buttons.

Phase 1 (this repository) is a VBA add-in (`Digit.xlam`). Presets are kept in a JSON file outside the add-in, and the formatting logic is kept apart from the UI, so a later C# COM/VSTO add-in can reuse both.

## Features

| | |
|---|---|
| **Quick buttons** | Up to 5 large icon buttons; hover one to see its group name. Their pictures and names come from the presets file. Slots with no group are hidden. |
| **Digit menu** | One button with every group that is not on a quick button, plus Settings, Reload presets, Check for updates and the version. **Undo Format** sits next to it. |
| **Cycle logic** | The active cell's current format picks the next preset. A format that matches no preset gets preset 1. |
| **Targeting** | Normal ranges and multi-area selections. A single cell in an Excel table formats the whole table column. A pivot values cell sets the DataField's number format, so it survives refresh. |
| **Undo Format** | Digit keeps its own history of the last 10 changes. Ctrl+Z works straight after a change. |
| **Shortcuts** | Ctrl+Alt+1 … Ctrl+Alt+5 by default. You can change them in Settings. |
| **Settings** | Create, rename, reorder and delete groups and presets. Includes a live preview, colour pickers, format checks, and import/export of presets. |
| **Updates** | **Check for updates** in the Digit menu, plus a quiet daily check when Excel starts. A new version is built and tested in the background and installed when you close Excel. See [docs/INSTALL.md](docs/INSTALL.md#updates). |
| **Storage** | `%APPDATA%\Digit\presets.json`. It survives add-in updates and can be shared with your team. |

## Quick start

On a Windows PC with Excel:

```powershell
git clone <this repo>
cd Digit
powershell -ExecutionPolicy Bypass -File build\build.ps1 -AllowVbomAccess -RunTests -Install
```

This builds `dist\Digit.xlam`, runs the self-tests in a hidden Excel, and installs the add-in for your user account. Start Excel and open the **Digit** tab. For details, other options, and Trust Center notes, see [docs/INSTALL.md](docs/INSTALL.md).

## Repository layout

```
src/
  ribbon/customUI14.xml           Ribbon definition (callbacks in DigitRibbon.bas)
  ribbon/images/                  Custom ribbon pictures (digitIcon.png), packed in by the build
  icons/                          Digit's number icons for quick buttons (and make_icons.py)
  vba/
    DigitCore.bas                 Pure logic: cycle rules, colours, shortcuts (no Excel objects)
    DigitFormat.bas               Engine: targeting, choosing the next preset, applying (no UI)
    DigitUndo.bas                 Undo snapshots and history (no UI)
    DigitIcons.bas                Digit's own button pictures: unpacks them and loads PNGs (GDI+)
    DigitUpdate.bas               Checks the release branch for updates and starts the installer (no UI)
    DigitStore.bas                presets.json: load, save, defaults, validation
    DigitRibbon.bas               Ribbon callbacks
    DigitApp.bas                  Start-up, shortcuts, messages, Ctrl+Z wiring
    DigitUI.bas                   Colour picker and other Windows helpers
    DigitTests.bas                Self-tests (unit and integration) that run inside Excel
    ThisWorkbook.code.vba         Workbook events, pasted into ThisWorkbook by the build
  forms/frmDigitSettings.code.vba Code for the Settings window
build/
  build.ps1                       Builds Digit.xlam through Excel COM automation
  DigitFormBuilder.bas            Lays out the Settings form at build time (not shipped)
  install-update.ps1              Builds, tests and installs an update in the background
lib/JsonConverter.bas             VBA-JSON 2.3.1 (MIT), Tim Hall
presets/presets.sample.json       The starter presets, written on first run
version.json                      Latest published version, read by Check for updates (on the release branch)
docs/                             Install, presets schema, manual build, architecture
```

## Documentation

- [docs/INSTALL.md](docs/INSTALL.md): build, install, Trust Center, signing, updating, uninstalling
- [docs/PRESETS.md](docs/PRESETS.md): the `presets.json` schema, and how cycling and targeting work
- [docs/MANUAL_BUILD.md](docs/MANUAL_BUILD.md): building by hand with the VBA editor and Office RibbonX Editor
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): module layering and notes for the C# port
- [docs/RELEASING.md](docs/RELEASING.md): publishing a new version so Digit can update itself

## License

The Digit sources belong to their authors. `lib/JsonConverter.bas` is VBA-JSON by Tim Hall, under the MIT licence (see `lib/JsonConverter.LICENSE`).
