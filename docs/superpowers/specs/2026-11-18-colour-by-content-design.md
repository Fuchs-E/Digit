# Colour by content: design

Date: 2026-11-18. Target version: 1.5.0.

## Intent

Financial models follow a font-colour convention so a reader can see at a glance
where a number comes from:

| Font colour | Means | Example |
|---|---|---|
| Blue | Hard-coded input typed into the cell | `1250`, `5.0%` |
| Black | Formula using only this sheet | `=B5*C5` |
| Green | Formula linking to another sheet in this workbook | `=Assumptions!C7` |
| Red | Formula linking to another file | `=[Budget.xlsx]Sheet1!A1` |

Today Digit can only apply one fixed font colour per preset, so colouring a
model means selecting inputs, formulas and links separately. The owner wants a
single click that colours any selection correctly, cell by cell.

## Goals

- One **Colour by content** button in the Digit group. Click it and every
  selected cell gets the font colour that matches what it contains.
- It changes **font colour only**. Number format, fill, bold / italic /
  underline, borders and everything else stay exactly as they are.
- The four colours are adjustable in Settings and stored in `presets.json`, so
  they are shared with the team like presets are.
- Undo Format and Ctrl+Z reverse it like any other Digit change.

## Non-goals

- No automatic colouring as you type, and no workbook-wide audit or report of
  miscoloured cells.
- No fill colours, and no separate colour for checks or errors.
- No change to how preset groups, cycling or quick buttons work.

## Anti-goals

- Touching anything other than font colour. A percentage that stops being a
  percentage is a failure.
- Colouring text labels and headers blue. They are typed constants, but they
  are not inputs.
- A whole-column or whole-sheet selection that hangs Excel.

## Approaches considered

1. **Its own fixed button in the Digit group (chosen).** Always present next to
   Undo Format, with its own shortcut and a colours section in Settings. Keeps
   the preset model ("number format + style") separate from this feature. Costs
   one more button of ribbon width.
2. **A special kind of preset group.** Could sit in any quick slot or the Digit
   menu. Rejected: it uses one of the five quick slots, and the preset editor
   would need to handle a group with no presets that does not cycle.
3. **Digit menu only.** Cheapest, but two clicks every time, against the
   "one click" goal.

The owner chose approach 1.

## Behaviour

### Classification (decided with the owner)

| Cell contains | Category | Default colour |
|---|---|---|
| Typed number, date, percentage, TRUE/FALSE | input | `#0000FF` |
| Formula made only of numbers and operators, e.g. `=1000*12` | input | `#0000FF` |
| Formula using only this sheet, e.g. `=B5*C5`, `=SUM(D2:D10)`, `=TODAY()` | formula | `#000000` |
| Formula with a typed number inside, e.g. `=B5*1.05` | formula (still a formula) | `#000000` |
| Formula pointing at another sheet in the same workbook | sheet link | `#008000` |
| Formula pointing at another workbook | external link | `#C00000` |
| A mix | strongest wins: external, then sheet link, then formula | |
| Typed text, blank cells | left alone | |

Further rules:

- **Errors.** A formula showing `#REF!` or similar is classified by its formula
  text like any other.
- **Own sheet.** `=Sheet1!A1` written on Sheet1 is a same-sheet formula (black).
- **3D references** such as `=SUM(Jan:Dec!B5)` are sheet links.
- **Named ranges and tables.** A name or structured table reference is coloured
  by where it actually points: a name that refers to another sheet makes the
  formula a sheet link, one that refers to another workbook makes it an
  external link, and one on the same sheet (or one that refers to a constant)
  leaves it a plain formula. Names that refer to other names are followed, to a
  small fixed depth.
- **Text inside quotes** is ignored: `="See Inputs!A1"` is a plain formula.
- **Formulas that return text**, e.g. `=A1&" units"`, are still formulas and are
  coloured. Only *typed* text is left alone.
- **Unknown words in a formula** (`LET` / `LAMBDA` variables, undefined names,
  names whose target is `#REF!`) contribute nothing to the classification. A
  sheet-scoped name wins over a workbook name of the same spelling, as in Excel.
- **Functions in another workbook or add-in**, e.g. `=PERSONAL.XLSB!MyFn(A1)`,
  are external links.
- **Values written by Excel rather than typed** (query or external-data output)
  are constants and are coloured as inputs.

### Targeting

- Works on the selection as it is: every cell of every area, on the active
  sheet. A single selected cell means that one cell. (Unlike preset buttons, a
  single cell in a table does **not** expand to the whole table column.)
- Only cells with content are touched, so whole-column and whole-sheet
  selections are limited to the used range.
- Cells inside a pivot table are skipped; Excel can wipe their fonts on refresh
  and they are neither inputs nor formulas.
- Merged cells are classified by their first cell.
- Charts, shapes, no workbook, protected sheets: same short messages as the
  preset buttons give today, no error dialogs.
- If nothing in the selection qualifies (all text or blank), nothing changes
  and the status bar says so.

### Feedback

Status bar after a run, for example:
`Digit: Colour by content  12 input, 30 formula, 4 sheet link, 1 external`.

The counts are cells classified into a category that has a colour set. Cells in
a category whose colour is blank are not counted. A run that applies anything
pushes an undo entry, even if every cell already had the right colour.

### Undo

One undo entry labelled "Colour by content", restoring the previous font colour
of every cell that was changed (including automatic and theme colours, which
the existing snapshot already handles). If the snapshot would exceed the
existing block limit, the existing "apply without undo?" question is asked.

The snapshot must be truly font-only. Today `Undo_CaptureRange` always records
and restores the number format as well, so `DigitUndo` gains a mode that leaves
number formats out. Otherwise undoing a colour run could revert a number format
the user changed in between, and models with varied formats would reach the
block limit sooner than they need to.

## Settings and storage

New keys under `settings` in `presets.json` (schema version stays 1; older
files get the defaults on load, and older Digit versions ignore the keys):

```json
"contentColors": {
  "input": "#0000FF",
  "formula": "#000000",
  "sheetLink": "#008000",
  "externalLink": "#C00000"
},
"contentColorShortcut": "Ctrl+Alt+6"
```

- Colours are `#RRGGBB`. A blank colour means "leave cells of this kind alone"
  (the same meaning a blank colour has in a preset). An invalid colour is
  repaired to that category's default.
- The shortcut follows the same rules and the same "Enable keyboard shortcuts"
  switch as the five quick-button shortcuts. Blank means none. It must not
  duplicate a quick-button shortcut; Settings refuses to save a duplicate. If a
  file is loaded where a quick button already uses the same shortcut (for
  example an older file that assigned Ctrl+Alt+6), the quick button keeps it
  and this shortcut is blanked.
- An older Digit version that saves Settings rewrites the file without these
  keys, so a shared file touched by a teammate who has not updated yet falls
  back to the default colours. The docs should say so.

**Settings window**, *Ribbon & shortcuts* tab: a new "Colour by content"
section with, for each of the four categories, a label, a colour swatch, a
**Choose...** button (the existing colour picker) and a **Clear** button (sets
the colour to blank, like the "none" buttons in the preset editor), plus a
**Reset to standard colours** button and the shortcut box. Reset changes the
four colours only, not the shortcut. Import with "replace all" takes the
imported file's colours and shortcut, as it already does for other settings.

## Ribbon

A new button `btnDigitColour` in the Digit group, placed directly after Undo
Format, sized by the same `Digit_GetSize` callback as the other buttons.
Screentip "Colour by content"; the supertip explains the four colours and shows
the active shortcut. Always visible and enabled.

## How it is built

Follows the existing layering in `docs/ARCHITECTURE.md`.

- **`DigitCore` (pure logic, no Excel objects).** A classifier that takes a
  formula's text (English A1 syntax, as `Range.Formula` returns it), the name
  of the sheet the cell is on, and a lookup of defined names and table names
  to where each points, and returns the category. Keeping this pure means it is
  unit-tested with plain strings and ports to the C# add-in unchanged.
- **`DigitFormat` (engine, no UI).** A new entry point alongside
  `Fmt_ApplyGroup`: validates the selection, builds the name lookup once for
  the workbook, finds the constant and formula cells in the selection,
  classifies them, snapshots the cells that will change with
  `Undo_CaptureRange` (font only), applies each colour to its cells in as few
  calls as practical, and reports a `DigitResult` and message.
- **`DigitStore`.** Defaults, normalisation/repair and accessors for the new
  keys.
- **`DigitApp` / `DigitRibbon` / Settings form / `DigitFormBuilder`.** Run
  entry point (button and shortcut share it, with the same reporting, undo
  arming and too-big-for-undo question as `App_RunGroup`), shortcut
  registration, ribbon callbacks, and the Settings section.
- **Docs and release.** README features table, `docs/PRESETS.md` (new keys and
  a "Colour by content" section), `docs/ARCHITECTURE.md` (new flow),
  `presets/presets.sample.json`, and the release steps in `docs/RELEASING.md`
  (`version.json`, `DIGIT_VERSION` in `DigitApp.bas`, release notes,
  `tools/check_vba.py`) for 1.5.0. Also correct the stale "Limits" line in
  `docs/PRESETS.md` that says theme colours are restored as RGB.

## Testing

In `DigitTests.bas`, run by `build.ps1 -RunTests`:

- **Classifier unit tests** covering every row of the classification table and
  each further rule, plus: absolute references, whole-column and whole-row
  references, function names that look like cell references (`LOG10(`),
  quoted sheet names with spaces and apostrophes, external links both open
  (`[Book.xlsx]Sheet1!A1`) and closed (`'C:\path\[Book.xlsx]Sheet1'!A1`),
  external names (`Book.xlsx!Name`), and structured references with brackets
  (`Table1[Amount]`, `Table1[[#This Row],[Amount]]`) which must **not** read as
  external.
- **Store tests**: defaults for a file without the keys, repair of invalid
  colours, round trip.
- **Integration tests** against the scratch workbook: a range mixing all
  categories with assorted number formats, fills and bold; after a run the font
  colours are right and number format / fill / bold are unchanged; text and
  blank cells untouched; a blank colour setting skips its category; undo
  restores every font colour; a whole-column selection completes and only
  touches used cells.

## Left to the builder

- The parsing technique inside the classifier, so long as the unit tests above
  pass.
- How cells are read and batched, so long as large selections stay
  responsive. Expected shape: read formulas in bulk, classify each distinct
  formula text once, and show progress in the status bar on long runs. The
  integration tests include one sizeable case.
- Exact layout of the new Settings section (the tab is currently full, so the
  form or existing frames may need resizing), and the ribbon button's picture.
- Behaviour of spilled dynamic-array cells: they should take the colour of the
  formula in the spill's anchor cell; confirm how Excel exposes them.

## Calls made by the designer (not explicitly decided by the owner)

- Typed TRUE/FALSE and typed error values count as inputs.
- `=TODAY()` and other formulas with functions but no references are plain
  formulas (black); only pure number-and-operator formulas count as inputs.
- A name that refers to a constant counts as a same-sheet reference.
- A single cell in a table colours that cell only, not the column.
- A blank colour means "skip this category".
- Default shortcut Ctrl+Alt+6; default red is dark red `#C00000`.
- A **Clear** button per colour in Settings, so "skip this category" does not
  need hand-editing the file.
- Undo for this feature never touches number formats.
- On a shortcut clash at load, the quick button keeps the shortcut.
- Functions in another workbook or add-in count as external links.
- Query and external-data output counts as inputs.
