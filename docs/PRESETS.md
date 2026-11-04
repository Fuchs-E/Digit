# Presets

Digit reads its presets from `%APPDATA%\Digit\presets.json`. The file is UTF-8 JSON. Both the VBA add-in and the planned C# add-in read it. You normally edit it in **Digit > Settings**, but it is fine to edit it by hand, keep it under version control, or share it with your team. After editing it by hand, click **Reload presets**.

Each time Digit saves the file, it keeps the previous version as `presets.json.bak`.

## Schema (version 1)

```json
{
  "version": 1,
  "settings": {
    "shortcutsEnabled": true,
    "shortcuts": ["Ctrl+Alt+1", "Ctrl+Alt+2", "Ctrl+Alt+3", "Ctrl+Alt+4", "Ctrl+Alt+5"],
    "undoDepth": 10
  },
  "slots": ["Numbers", "Thousands", "Percent", "Currency", ""],
  "groups": [
    {
      "name": "Numbers",
      "presets": [
        { "numberFormat": "#,##0;(#,##0);-",       "fontColor": "",        "fillColor": "" },
        { "numberFormat": "#,##0.00;(#,##0.00);-", "fontColor": "#0000FF", "fillColor": "#FFF2CC" }
      ]
    }
  ]
}
```

| Field | Type | Meaning |
|---|---|---|
| `version` | number | Schema version. Currently `1`. |
| `settings.shortcutsEnabled` | bool | Turns the quick-button keyboard shortcuts on or off. |
| `settings.shortcuts` | 5 strings | The shortcut for buttons 1 to 5, such as `"Ctrl+Alt+1"`, `"Ctrl+Shift+F5"` or `"Alt+Q"`. Use `""` for none. A shortcut needs Ctrl or Alt, unless it is a function key. |
| `settings.undoDepth` | number | How many changes Undo Format remembers (1 to 50, default 10). |
| `slots` | 5 strings | The group name on quick buttons 1 to 5, or `""`. Buttons with no group are hidden. Groups that are not in a slot appear under **More presets**. |
| `groups[].name` | string | Must be unique (case does not matter) and not empty. It is shown as the button label. |
| `groups[].presets` | 1 to 3 objects | The cycle order. |
| `presets[].numberFormat` | string | **Required.** An Excel number format code, written in the English (US) syntax that VBA's `Range.NumberFormat` uses, for example `#,##0,;(#,##0,);-` or `0.0%`. The file then works the same in every Windows locale. |
| `presets[].fontColor` | string | `"#RRGGBB"`, or `""` to leave the font colour as it is. |
| `presets[].fillColor` | string | `"#RRGGBB"`, or `""` to leave the fill as it is. |

Digit repairs the file when it loads it rather than refusing it:
- Unknown keys are ignored.
- Missing keys get their defaults.
- Groups with no name or no presets are skipped.
- Duplicate group names get a suffix, such as `" (2)"`.
- Presets after the third are dropped.
- Invalid colours become blank.
- Slots that name a missing group are cleared.

If the file is not valid JSON at all, Digit renames it to `presets.invalid-<timestamp>.json` and writes the starter presets in its place.

## How a click is resolved

1. **What to read.** Digit reads the number format, font colour and fill colour of the **active cell**. If the active cell is merged, it reads the merged area's first cell. When a selection has mixed formats, the active cell decides.
2. **Which preset it matches.** Digit compares the active cell with the group's presets in order. The number format must match. Digit also accepts Excel's stored spelling of a format: Excel sometimes rewrites a format code, for example `$` as `\$`. If a preset also sets a font or fill colour, that colour must match too. A blank colour matches anything. When several presets match, the one that sets the most colours wins.
3. **What to apply next.** No match applies preset 1. A match on preset *k* applies preset *k+1*. A match on the last preset applies preset 1 again. A group with fewer than three presets only cycles through the ones it has.
4. **What gets formatted.** The chosen preset is applied to the **whole target in one go**:
   - **Normal selection**: every cell of every area.
   - **One cell in an Excel table's data body**: that table column's whole `DataBodyRange`. If you select several cells in a table, only the selection is formatted.
   - **A cell in a pivot table's values area** (values, subtotals, grand totals, or a data field header): Digit sets `PivotField.NumberFormat` on that cell's data field. If the selection covers several data fields, it sets each one. On pivots, only the number format is applied: font and fill colours are skipped, because Excel can wipe them on refresh.
   - **Charts, shapes, or no selection**: nothing happens, and the status bar says why.
   - **Protected sheets**: you see a short message instead of an error. Sheets protected with *Format cells* allowed, or with `UserInterfaceOnly`, still work.

### Tip: presets that share a number format

A blank colour means "leave as is". Digit never removes a colour. Suppose preset 1 is `#,##0` with a blue font, and preset 2 is the same format with a blue font and a yellow fill. Going from preset 2 back to preset 1 leaves the yellow fill in place, so the cell keeps matching preset 2. If presets in one group share a number format, give each of them explicit font and fill colours, for example `#FFFFFF` for a white fill instead of blank. Or give each preset a different number format.

## Undo

Before each change, Digit saves what it is about to overwrite: the number format, plus the font colour and fill only when the preset changes them. For pivots, it saves the data field formats. **Undo Format** restores the last change. Ctrl+Z does the same straight after a Digit change, as long as you have done nothing else in Excel since. The history holds the last `undoDepth` changes for the current Excel session.

Digit saves undo data per uniform block of cells, not per cell, so large and whole-column selections stay fast. If a selection is so varied that saving it would take too long, Digit asks before applying the format without undo.

Limits: Digit restores theme colours as plain RGB colours, and it restores gradient or patterned fills as solid fills.
