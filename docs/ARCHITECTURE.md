# Architecture

## Layers

```
 customUI14.xml ──► DigitRibbon ─┐
 OnKey shortcuts ──► DigitApp ◄──┤                     UI layer
 frmDigitSettings ──────────────┘   (MsgBox, status bar, form, ribbon handle)
                        │
                        ▼
          DigitFormat ──► DigitUndo                    Engine (Excel object model, no UI)
                        │
                        ▼
          DigitStore ──► JsonConverter                 Storage (presets.json)
                        │
                        ▼
                    DigitCore                          Pure logic (no Excel objects)
```

| Module | Knows about | Must not know about |
|---|---|---|
| `DigitCore` | Strings, numbers, collections | Excel, UI |
| `DigitStore` | The JSON file and the config tree | Ranges, UI |
| `DigitFormat` / `DigitUndo` | `Range`, `ListObject`, `PivotTable` | MsgBox, ribbon, forms |
| `DigitApp`, `DigitRibbon`, `DigitUI`, the form | Everything above | |

The engine reports back with a `DigitResult` code and a message (`drOK`, `drNothingToDo`, `drWarning`, `drUndoTooBig`, `drError`). `DigitApp` decides whether that becomes a status-bar note or a message box, and it asks the user when undo would be too large. This split is what lets `DigitTests` drive the engine without dialogs.

## Key flows

**Click a quick button:** `Digit_OnQuick` (tag = slot) → `App_RunSlot` → `App_RunGroup`. That calls `Fmt_ApplyGroup`, which does the following:
1. Resolves the target: a pivot data field, a table column, or the selection.
2. Reads the active cell's state.
3. Uses `Core_BestMatch` and `Core_NextIndex` to pick the next preset.
4. Snapshots the target with `Undo_CaptureRange` or `Undo_CapturePivot`.
5. Applies the preset in one call per property.
6. Pushes the snapshot onto the undo history.

`DigitApp` then arms `Application.OnUndo` for Ctrl+Z and refreshes the Undo button.

**Presets changed** (Settings > Save, Reload, or import): `App_OnPresetsChanged` clears the format cache, re-registers the shortcuts, and calls `IRibbonUI.Invalidate`. Labels, visibility and the More menu all update through the callbacks.

**Loss of VBA state:** an unhandled error or a VBA reset clears module variables. `Store_Config()` reloads lazily. The ribbon handle is recovered from the pointer saved on the add-in's hidden sheet. Every entry point calls `Digit_Startup`, which only runs once.

## Notes for the C# COM/VSTO port

The port should keep these unchanged:
- `customUI14.xml`: the same control ids and callback names, mapped to C# methods with the same signatures.
- `presets.json`: the schema in [PRESETS.md](PRESETS.md). Colours are `#RRGGBB` strings, shortcuts are readable text, and number formats use English codes, so nothing in the file is VBA-specific.
- The cycle rules (`Core_BestMatch`, `Core_NextIndex`) and targeting (`Fmt_RangeTarget`, `Fmt_PivotTarget`). They translate method for method.

Differences to plan for:

| VBA | C# |
|---|---|
| `Application.OnKey` | No equivalent for COM add-ins. Use a low-level keyboard hook, or VSTO's `Application.OnKey` with a VBA stub. |
| `Application.OnUndo` | Same limitation. Keep Digit's own history. A COM add-in can still call `Application.OnUndo` with a macro in a helper `.xlam`. |
| Scratch cell for format checks and previews | Same approach, using a hidden worksheet in a temporary workbook, or Excel's `TEXT` worksheet function. |
| `Scripting.Dictionary` tree | Typed classes (`DigitConfig`, `PresetGroup`, `Preset`) with System.Text.Json. |
| Ribbon pointer hack | Not needed. The ribbon object lives as long as the add-in. |
| Colour Longs (BGR) | `System.Drawing.ColorTranslator.ToOle` / `FromOle`. |

The integration tests in `DigitTests.bas` describe the expected behaviour precisely. Port them as the C# acceptance tests.
