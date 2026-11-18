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
    "undoDepth": 10,
    "contentColors": {
      "input": "#0000FF",
      "formula": "#000000",
      "sheetLink": "#008000",
      "externalLink": "#C00000"
    },
    "contentColorShortcut": "Ctrl+Alt+6",
    "contentColorSlot": 0
  },
  "slots": ["Numbers", "Thousands", "Percent", "Currency", ""],
  "groups": [
    {
      "name": "Numbers",
      "icon": "CommaStyle",
      "presets": [
        { "numberFormat": "#,##0;(#,##0);-",       "fontColor": "",        "fillColor": "" },
        { "numberFormat": "#,##0.00;(#,##0.00);-", "fontColor": "#0000FF", "fillColor": "#FFF2CC", "bold": true }
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
| `settings.largeButtons` | bool | `true` shows the quick buttons and Undo Format large on the ribbon (one per column); `false` (default) shows them small, three to a column. Set in **Settings > Large quick buttons**. |
| `settings.undoDepth` | number | How many changes Undo Format remembers (1 to 50, default 10). |
| `settings.contentColors` | object | The font colours **Colour by content** applies: `input`, `formula`, `sheetLink` and `externalLink`, each `"#RRGGBB"`, or `""` to leave cells of that kind alone. See [Colour by content](#colour-by-content). |
| `settings.contentColorShortcut` | string | The shortcut for Colour by content (default `"Ctrl+Alt+6"`, `""` for none). It follows `shortcutsEnabled`. |
| `settings.contentColorSlot` | number | Where the Colour by content button sits. `0` (default): its own button after Undo Format. `1` to `5`: in place of that quick button. Four preset groups then stay on quick buttons, around it in list order, and the rest are in the Digit menu. That position's shortcut in `settings.shortcuts` runs Colour by content too. |
| `slots` | 5 strings | Written by Digit for reference: the names of the groups on the quick buttons (`""` at a position Colour by content holds). **The quick buttons are the first five groups in `groups`, left to right**; the rest appear in the Digit menu. When Digit reads a file, groups named in `slots` move to the front in that order, so files from older versions keep their buttons. |
| `groups[].name` | string | Must be unique (case does not matter) and not empty. It is shown as the button label. |
| `groups[].icon` | string | Optional. The picture on the group's quick button and in the Digit menu. Either one of Digit's own icons, written `"digit:<name>"` (for example `"digit:pct_green"`; the names are the files in `src/icons`), or a built-in Office icon name (an `imageMso` name) such as `"PercentStyle"`. `""` or an unknown name shows the button number instead. Pick one in **Settings > Button picture > Choose...**, which shows the pictures themselves, a page at a time. |
| `groups[].presets` | 1 to 3 objects | The cycle order. |
| `presets[].numberFormat` | string | **Required.** An Excel number format code, written in the English (US) syntax that VBA's `Range.NumberFormat` uses, for example `#,##0,;(#,##0,);-` or `0.0%`. The file then works the same in every Windows locale. |
| `presets[].fontColor` | string | `"#RRGGBB"`, or `""` to leave the font colour as it is. |
| `presets[].bold`, `presets[].italic`, `presets[].underline` | bool | Optional. `true` turns the style on, `false` turns it off. Leave the field out to leave the cell's style as it is. Like colours, these also tell apart presets that share a number format. |
| `presets[].fillColor` | string | `"#RRGGBB"`, or `""` to leave the fill as it is. |

Digit repairs the file when it loads it rather than refusing it:
- Unknown keys are ignored.
- Missing keys get their defaults.
- Groups with no name or no presets are skipped.
- Duplicate group names get a suffix, such as `" (2)"`.
- Presets after the third are dropped.
- Invalid colours become blank.
- A missing or invalid Colour by content colour gets its standard colour.
- A Colour by content shortcut that a quick button already uses is cleared.
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

## Colour by content

The **Colour by content** button (next to Undo Format, or in place of a quick button if you choose that in **Settings > Colour by content > Position on the ribbon**; default shortcut Ctrl+Alt+6) sets the font colour of every selected cell by what the cell contains. It is the standard financial-modelling convention:

| Cell contains | Kind | Default colour |
|---|---|---|
| A typed number, date, percentage, TRUE/FALSE or error value | `input` | Blue `#0000FF` |
| A formula made only of numbers and operators, such as `=1000*12` | `input` | Blue `#0000FF` |
| A formula that stays on its own sheet, such as `=B5*C5`, `=SUM(D2:D10)`, `=TODAY()` or `=B5*1.05` | `formula` | Black `#000000` |
| A formula that points at another sheet in the workbook, such as `=Assumptions!C7` | `sheetLink` | Green `#008000` |
| A formula that points at another file, such as `=[Budget.xlsx]Sheet1!A1` | `externalLink` | Dark red `#C00000` |

**Only the font colour changes.** The number format, fill, bold / italic / underline and borders stay as they are.

Details:
- **Mixed formulas** take the strongest kind: external link, then sheet link, then formula. `=B5*Inputs!C2` is a sheet link.
- **Named ranges and tables** count by where they point. `=Revenue` is a sheet link when `Revenue` refers to another sheet, and a plain formula when it refers to the same sheet or to a constant. Names that refer to other names are followed.
- **Text in quotes is not a reference.** `="See Inputs!A1"` is a plain formula.
- **A formula naming its own sheet**, such as `=Model!A1` on the sheet Model, is a plain formula.
- **Functions from another workbook or add-in**, such as `=PERSONAL.XLSB!MyFn(A1)`, are external links.
- **Typed text and blank cells are left alone**, so labels and headers keep their colour. A formula that returns text is still a formula and is coloured.
- **Pivot tables are skipped.**
- **Values written by Excel**, such as the output of a query, are constants and are coloured as inputs.
- **Dynamic array formulas** colour the cells they spill into as well (when the selection holds up to 5,000 formula cells).
- **The selection is used as it is.** One selected cell colours that cell only, including inside a table. Whole columns, rows or sheets are fine: Digit only looks at cells with content.
- **A blank colour** means "leave cells of this kind alone". Set it with **Clear** in **Settings > Colour by content**.
- **Protected sheets** must be unprotected first.

The status bar shows what was coloured, for example `Digit: Colour by content  12 input, 30 formula, 4 sheet link, 1 external`. **Undo Format** (or Ctrl+Z straight afterwards) puts the previous font colours back and touches nothing else.

**Sharing the colours.** They live in `presets.json` with your presets. Digit 1.4 and earlier ignore these keys and drop them when they save the file, so a shared file saved by someone who has not updated yet goes back to the standard colours.

## Undo

Before each change, Digit saves what it is about to overwrite: the number format, plus the font colour and fill only when the preset changes them. For pivots, it saves the data field formats. For Colour by content, it saves the font colours only. **Undo Format** restores the last change. Ctrl+Z does the same straight after a Digit change, as long as you have done nothing else in Excel since. The history holds the last `undoDepth` changes for the current Excel session.

Digit saves undo data per uniform block of cells, not per cell, so large and whole-column selections stay fast. If a selection is so varied that saving it would take too long, Digit asks before applying the format without undo.

Limits: Digit restores gradient or patterned fills as solid fills. A cell whose text mixes several font colours does not get them back.
