# Colour by content Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A one-click **Colour by content** button that sets each selected cell's font colour by what it contains (input, formula, sheet link, external link) and changes nothing else.

**Architecture:** A pure classifier in `DigitCore` turns a formula's text into a kind. A new engine entry point in `DigitFormat` finds the constant and formula cells in the selection, classifies them, snapshots font colours through a new font-only mode in `DigitUndo`, and applies one colour per kind. Colours and the shortcut live under `settings` in `presets.json` (`DigitStore`), edited on a new third Settings tab.

**Tech Stack:** Excel VBA (`.bas` sources built into `Digit.xlam` by `build/build.ps1` on Windows), Python for `tools/check_vba.py` and the ribbon icon.

**Spec:** `docs/superpowers/specs/2026-11-18-colour-by-content-design.md`

## Global Constraints

- Only font colour changes. Number format, fill, bold / italic / underline are never written.
- Default colours: input `#0000FF`, formula `#000000`, sheetLink `#008000`, externalLink `#C00000`. Default shortcut `Ctrl+Alt+6`.
- `presets.json` schema version stays `1`.
- Layering from `docs/ARCHITECTURE.md`: `DigitCore` has no Excel objects; `DigitFormat` / `DigitUndo` have no UI.
- Excel is not available in the development container. Each task is verified with `python3 tools/check_vba.py` (must print `0 problem(s)`). The VBA self-tests added here run on Windows with `build\build.ps1 -AllowVbomAccess -RunTests`.
- The classifier algorithm is first proven in Python (scratch, not committed) against the same cases as the VBA unit test, then transcribed to VBA.

## Review Focus

1. A single selected cell: `SpecialCells` on one cell searches the whole sheet, so a one-cell selection must not colour other cells. (Task 4 test: single cell.)
2. Structured table references (`Table1[Amount]`, `[@Amount]`) must not read as external links. (Task 1 tests.)
3. A selection of only text or blank cells changes nothing and pushes no undo entry. (Task 4 test.)
4. Undo after the user changed a number format in between must not revert that format. (Task 4 test.)
5. A shortcut clash between Colour by content and a quick button on load. (Task 2 test.)

---

### Task 1: Classifier in DigitCore

**Files:** Modify `src/vba/DigitCore.bas`, `src/vba/DigitTests.bas`.

**Interfaces — Produces:**

```vba
Public Enum DigitContentKind
    ckNone = 0          ' not a formula / leave alone
    ckInput = 1
    ckFormula = 2
    ckSheetLink = 3
    ckExternal = 4
End Enum
Public Const DIGIT_CONTENT_KINDS As Long = 4

' formula: text starting with "=" (A1 or R1C1, English). ownSheet: the cell's sheet name.
' names: LCase(name) -> RefersTo text; sheet-scoped names keyed LCase(sheet) & "!" & LCase(name);
'        tables keyed LCase(table name) -> "='Sheet'!A1". May be Nothing.
' sheets: LCase(sheet name) -> anything, for every sheet in the workbook. May be Nothing
'         (then any unbracketed qualifier counts as a sheet).
Public Function Core_ClassifyFormula(ByVal formula As String, ByVal ownSheet As String, _
        ByVal names As Scripting.Dictionary, ByVal sheets As Scripting.Dictionary) As DigitContentKind
```

Rules the scanner implements (text inside `"..."` is skipped):
- `'...'!` qualifier: contains `[` → external; else its sheet (the part before `:` for 3D) not in `sheets` → external; else equal to `ownSheet` (and not 3D) → no link; else sheet link.
- `[...]` (nesting, `'` escapes the next character inside): followed by a word and `!` → external; otherwise a structured or R1C1 reference → counts as a reference.
- word followed by `!`, or `word:word!` (3D): same qualifier rule as above without the bracket case.
- word followed by `(` → function. `TRUE` / `FALSE` → literal. A word found in `names` (own-sheet key first, then global) → classify its target recursively, depth limit 5; an input / none result counts as formula. Any other word → formula.
- `"..."` text literal → formula. `:` outside brackets → reference. `#...` error literal → formula. Numbers (with exponent) → nothing.
- Result: external > sheet link > formula; nothing but numbers and operators → input (a text literal makes it a formula). Text not starting with `=` → `ckNone`.

- [ ] Step 1: Write `Test_Classify` in `DigitTests.bas` and call it from `RunAllTests` after `Test_Collections`. Cases (own sheet `Model`; sheets `model`, `inputs`, `my sheet`, `jan`, `dec`; names `taxrate`→`=0.05`, `revenue`→`=Inputs!$C$7`, `local`→`=Model!$A$1`, `ext`→`='[Budget.xlsx]Sheet1'!$A$1`, `chain`→`=Revenue*2`, `model!scoped`→`=Inputs!$A$1`, `scoped`→`=Model!$A$1`, `table1`→`='Inputs'!A1`, `table2`→`='Model'!A1`, `loop`→`=Loop+1`):

| Formula | Expected |
|---|---|
| `1250` | ckNone |
| `=1000*12`, `=-(5+3)/2`, `=1E+3*2`, `=TRUE` | ckInput |
| `=B5*C5`, `="a"&"b"`, `=SUM(D2:D10)`, `=TODAY()`, `=B5*1.05`, `=$A$1`, `=SUM(A:A)`, `=SUM(1:1)`, `=LOG10(100)`, `=Model!A1`, `="See Inputs!A1"`, `=A1&" units"`, `=#REF!+1`, `=TaxRate*B5`, `=Local`, `=Table2[Amount]`, `=[@Amount]*2`, `=Table2[[#This Row],[Amount]]`, `=LET(x,1,x*2)`, `=Undefined+1`, `=RC[-1]*R[-1]C`, `=A1#`, `=Loop` | ckFormula |
| `=Assumptions!C7` with `sheets` Nothing, `=Inputs!C7`, `='My Sheet'!A1`, `=B5*Inputs!C2`, `=SUM(Jan:Dec!B5)`, `=SUM('Jan:Dec'!B5)`, `=Revenue`, `=Chain`, `=Scoped`, `=Table1[Amount]`, `=SUM(Table1[[#Totals],[Amount]])` | ckSheetLink |
| `=[Budget.xlsx]Sheet1!A1`, `=[1]Sheet1!A1`, `='C:\path\[Budget.xlsx]Sheet1'!A1`, `=Budget.xlsx!Name`, `='C:\x\Budget.xlsx'!Name`, `=PERSONAL.XLSB!MyFn(A1)`, `=Inputs!A1+[Budget.xlsx]Sheet1!A1`, `=Ext` | ckExternal |

- [ ] Step 2: Prove the algorithm in Python against the table (scratch file).
- [ ] Step 3: Implement the enum, constant and `Core_ClassifyFormula` in `DigitCore.bas`.
- [ ] Step 4: `python3 tools/check_vba.py` → `0 problem(s)`.
- [ ] Step 5: Commit.

### Task 2: Settings in DigitStore

**Files:** Modify `src/vba/DigitStore.bas`, `src/vba/DigitTests.bas`, `presets/presets.sample.json`.

**Interfaces — Consumes:** `DigitContentKind`. **Produces:**

```vba
Public Function Store_ContentKey(ByVal kind As DigitContentKind) As String       ' "input", "formula", "sheetLink", "externalLink"
Public Function Store_ContentDefault(ByVal kind As DigitContentKind) As String   ' "#0000FF", "#000000", "#008000", "#C00000"
Public Function Store_ContentColor(ByVal cfg As Scripting.Dictionary, ByVal kind As DigitContentKind) As String
Public Sub Store_SetContentColor(ByVal cfg As Scripting.Dictionary, ByVal kind As DigitContentKind, ByVal hexColor As String)
Public Function Store_ContentShortcut(ByVal cfg As Scripting.Dictionary) As String
Public Sub Store_SetContentShortcut(ByVal cfg As Scripting.Dictionary, ByVal shortcut As String)
```

`Store_DefaultConfig` and `Store_Normalize` write `settings("contentColors")` (Dictionary with the four keys) and `settings("contentColorShortcut")`. Normalise: key missing → default; `""` stays `""`; invalid → default; shortcut missing → `Ctrl+Alt+6`; shortcut whose `Core_ShortcutToOnKey` equals a quick-button shortcut's → `""`.

- [ ] Step 1: Add `Test_ContentSettings` (called after `Test_Normalize`): empty `{}` config gives the four defaults and `Ctrl+Alt+6`; `{"settings":{"contentColors":{"input":"","formula":"zzz","sheetLink":"#00ff00"}}}` gives `""`, `#000000`, `#00FF00`, `#C00000`; shortcuts `["Ctrl+Alt+6","","","",""]` with no `contentColorShortcut` gives `""`; `"contentColorShortcut":"ctrl+shift+k"` gives `Ctrl+Shift+K`; `Store_Clone` keeps a set colour and shortcut.
- [ ] Step 2: Implement; add the two keys to `presets/presets.sample.json`.
- [ ] Step 3: `python3 tools/check_vba.py` → `0 problem(s)`. Commit.

### Task 3: Font-only undo snapshot

**Files:** Modify `src/vba/DigitUndo.bas`.

**Interfaces — Produces:**

```vba
' Snapshot of the font colour only (no number format) of every range in parts
' (Collection of Range, all on one sheet). Nothing when over MAX_BLOCKS or parts is empty.
Public Function Undo_CaptureFont(ByVal parts As Collection, ByVal label As String) As Scripting.Dictionary
```

Blocks from this capture hold `Empty` in the number-format slot; `ReadUniform` does not split on number format in this mode; `Undo_Restore` writes `NumberFormat` only when the slot is not `Empty`. Tested through Task 4's integration test.

- [ ] Step 1: Implement. `python3 tools/check_vba.py` → `0 problem(s)`. Commit.

### Task 4: Engine entry point

**Files:** Modify `src/vba/DigitFormat.bas`, `src/vba/DigitTests.bas`.

**Interfaces — Consumes:** Tasks 1–3. **Produces:**

```vba
' Colours the selection's cells by content. Same result codes and skipUndo meaning as Fmt_ApplyGroup.
Public Function Fmt_ColourByContent(ByRef msg As String, Optional ByVal skipUndo As Boolean = False) As DigitResult
```

Behaviour: validates the selection with the same messages as `Fmt_ApplyGroup`; limits work to `Intersect(Selection, UsedRange)`; constants are numbers, logicals and errors (`SpecialCells`, with a direct test when the target is one cell); formulas are read per area as `FormulaR1C1` arrays and each distinct text classified once; cells inside any pivot table's `TableRange2` are skipped; kinds whose colour is blank are skipped; spill ranges take their anchor's kind when there are at most 5,000 formula cells. Pieces are collected per kind in chunks of at most 500 areas. Undo label `Colour by content`. Messages: `Digit: Colour by content  12 input, 30 formula, 4 sheet link, 1 external` (zero counts left out) and `Digit: nothing to colour in the selection (no numbers or formulas).` with `drNothingToDo`.

- [ ] Step 1: Add `Test_ColourByContent` (called after `Test_Styles`). Scratch book with sheets `Model`, `Inputs`; on `Model`: A1 `1250` format `0.0%` bold with yellow fill, A2 `="x"` … covering each kind, text `"Label"` in A7 with a red font, blank A8, a name `Revenue`→`Inputs!$C$7`. Assert after a run on `A1:B10`: each font colour; A1 number format, fill and bold unchanged; A7 still red; `Undo_Count` = 1; change A1's number format to `0.00`, undo, A1 font back to its old colour and number format still `0.00`. Single cell: select only A1 after resetting colours, run, B-column cells unchanged. Text-only selection: `drNothingToDo`, no undo entry. Blank input colour in the temp presets: constants untouched. Whole column `A:A`: `drOK`. Shape selected: `drNothingToDo`. Protected sheet: `drWarning`.
- [ ] Step 2: Implement. `python3 tools/check_vba.py` → `0 problem(s)`. Commit.

### Task 5: Button, shortcut and ribbon

**Files:** Modify `src/vba/DigitApp.bas`, `src/vba/DigitRibbon.bas`, `src/ribbon/customUI14.xml`. Create `src/ribbon/images/make_colourIcon.py`, `src/ribbon/images/colourIcon.png`.

**Interfaces — Produces:**

```vba
Public Sub App_RunColour()                       ' DigitApp: run, report, undo-too-big question, arm Ctrl+Z
Public Sub DigitColourByContent()                ' OnKey target
Public Function App_ColourShortcut() As String   ' active shortcut text or ""
Public Sub Digit_OnColour(ByVal control As IRibbonControl)
Public Sub Digit_GetColourScreentip(ByVal control As IRibbonControl, ByRef returnedVal)
Public Sub Digit_GetColourSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
```

Ribbon: `<button id="btnDigitColour" label="Colour by content" showLabel="false" getSize="Digit_GetSize" image="colourIcon" .../>` directly after `btnDigitUndo`. The shortcut registers under the same `shortcutsEnabled` switch.

- [ ] Step 1: Implement and draw the icon. Check the XML parses (`python3 -c "import xml.dom.minidom,sys; xml.dom.minidom.parse('src/ribbon/customUI14.xml')"`). `python3 tools/check_vba.py` → `0 problem(s)`. Commit.

### Task 6: Settings tab

**Files:** Modify `build/DigitFormBuilder.bas`, `src/forms/frmDigitSettings.code.vba`, `src/vba/DigitTests.bas`.

A third tab, **Colour by content**, because the *Ribbon & shortcuts* tab is full: four rows (`lblCcName1..4`, `lblCcSwatch1..4`, `btnCcPick1..4`, `btnCcNone1..4`), `btnCcReset` ("Reset to standard colours", colours only), `txtKeyColour`. Save validates the shortcut and refuses a duplicate of a quick-button shortcut. `Test_SettingsForm` expects three tabs and a swatch control.

- [ ] Step 1: Implement. `python3 tools/check_vba.py` → `0 problem(s)`. Commit.

### Task 7: Docs and version 1.5.0

**Files:** Modify `README.md`, `docs/PRESETS.md`, `docs/ARCHITECTURE.md`, `version.json`, `src/vba/DigitApp.bas` (`DIGIT_VERSION`).

- [ ] Step 1: Update docs (features row, new keys, a "Colour by content" section, the older-version caveat, the stale theme-colour "Limits" line, the new flow). Bump the version and notes. `python3 tools/check_vba.py` → `0 problem(s)`. Commit.
