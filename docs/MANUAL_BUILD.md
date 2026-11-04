# Building Digit by hand

Use these steps if you can't run `build\build.ps1`, for example because PowerShell scripts are blocked. You need Excel and the free [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).

## 1. Create the workbook and import the code

1. In Excel, create a new blank workbook. Delete all sheets but one, and rename that sheet `Digit`.
2. Press **Alt+F11** to open the VBA editor.
3. Go to **Tools > References** and tick **Microsoft Scripting Runtime**. VBA-JSON needs it.
4. **File > Import File...** each of these:
   - `lib\JsonConverter.bas`
   - every `src\vba\*.bas` (`DigitTests.bas` is optional)
5. Open the **ThisWorkbook** module (it may have a localised name, such as *DieseArbeitsmappe*). Paste in the contents of `src\vba\ThisWorkbook.code.vba`.

## 2. Create the Settings form

The form's layout is generated from code, so there is no binary `.frx` to keep in the repository.

1. Turn on **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**.
2. Import `build\DigitFormBuilder.bas`.
3. In the Immediate window (**Ctrl+G**), run:
   ```
   DigitFormBuilder.BuildSettingsForm "C:\path\to\Digit\src\forms\frmDigitSettings.code.vba"
   ```
4. Remove the `DigitFormBuilder` module: right-click it, choose **Remove**, and answer **No** when asked to export it.
5. Run **Debug > Compile VBAProject**. It should finish without errors.
6. Optional: run `Digit_RunTests` from the Immediate window (or with Alt+F8). It shows the self-test results.

If someone has already built the add-in with `-ExportForm`, you can instead **File > Import** `dist\form\frmDigitSettings.frm` (keep the `.frx` next to it).

## 3. Save as an add-in

**File > Save As**, choose **Excel Add-in (\*.xlam)**, name it `Digit.xlam`, and close Excel.

## 4. Add the ribbon with Office RibbonX Editor

1. Open `Digit.xlam` in Office RibbonX Editor.
2. Go to **Insert > Office 2010+ Custom UI Part**. This creates `customUI14.xml`.
3. Replace its contents with `src\ribbon\customUI14.xml`.
4. Click **Validate**, then **Save**.

The editor adds the package relationship for you: type `http://schemas.microsoft.com/office/2007/relationships/ui/extensibility`, target `customUI/customUI14.xml`. `build.ps1` adds the same relationship.

Now install the add-in as described in [INSTALL.md](INSTALL.md#2-install-in-excel).
