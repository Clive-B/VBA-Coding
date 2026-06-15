# VBA Coding Handoff

Last updated: 2026-06-15

Repository: https://github.com/Clive-B/VBA-Coding.git

Workspace:

```text
C:\Users\codro\Desktop\VBA Coding
```

## Operating Rules

- Make code changes in the exported `.bas` files in this folder.
- Do not import/update modules inside Excel workbooks unless the user explicitly asks for workbook updates.
- When a new file appears in this folder, add it to the repo and push it.
- The current Git branch is `main`.
- Git sometimes shows warnings about `C:\Users\codro/.config/git/ignore` permission; those warnings have not blocked commits.

## Current Module Files

```text
CovFiles.bas
CSTCOVMOSSingleMacro.bas
CSTFiles.bas
CSTSingleFileProcessing.bas
DataFiles.bas
FourGCovAUI.bas
FourGDataMacro.bas
OpenCSTForm.bas
Spots.bas
ThreeGCovUI.bas
ThreeGDataMacro.bas
```

## Important CST Status

`CSTCOVMOSMacro.bas` was renamed to:

```text
CSTCOVMOSSingleMacro.bas
```

The VBA module attribute inside the file was also changed to:

```vba
Attribute VB_Name = "CSTCOVMOSSingleMacro"
```

The public macro procedure name remains:

```vba
Public Sub CSTCOVMOS(...)
```

This was intentional so existing `Application.Run "...!CSTCOVMOS"` calls can continue to work.

## Single CST File Processing

A new module was added:

```text
CSTSingleFileProcessing.bas
```

Purpose:

- Support the new UserForm-driven CST workflow.
- User selects one `.nmf` or `.nmfs` log file.
- The selected file can be either MOC or MTC.
- The code uses that selected file as an anchor.
- It searches only the selected file's folder for related CST logs.
- It groups matching split files for the same CST session/location.
- It requires both MOC and MTC logs before processing.
- It opens `Voice Template.xlsm` and runs `CSTCOVMOS` only for the related log set.

Main entry points:

```vba
Public Sub ProcessSingleCSTFromForm()
Public Sub ProcessSingleCSTFromSubmittedForm(Optional ByVal frm As Object = Nothing)
Public Sub ProcessSingleCSTFiles(ByVal cstExcelName As String, ByVal selectedCSTFilePath As String, ByVal storageFolderPath As String)
```

The UserForm Send button should call:

```vba
FormSubmitted = True
Me.Hide
ProcessSingleCSTFromSubmittedForm Me
```

If the form is launched through a standard macro instead, assign the button/menu to one of these launchers in `OpenCSTForm.bas`:

```vba
OpenQoSTemplateForm
OpenCSTForm
OpenSingleCSTForm
```

All three call:

```vba
ProcessSingleCSTFromForm
```

## Expected UserForm Public Fields

The UserForm should expose these public variables:

```vba
Public cstExcelName As String
Public selectedCSTFilePath As String
Public storageFolderPath As String
Public FormSubmitted As Boolean
```

The file picker should allow:

```vba
*.nmf; *.nmfs
```

Recommended UserForm messages:

- File picker title: `Select MOC or MTC NMF Log`
- Missing file message: `Please select the CST log file.`

## CST Output Override

`CSTCOVMOSSingleMacro.bas` now includes optional single-run output controls:

```vba
Public SingleCSTOutputEnabled As Boolean
Public SingleCSTOutputFolder As String
Public SingleCSTOutputName As String

Public Sub ConfigureSingleCSTOutput(ByVal outputFolder As String, ByVal outputName As String, Optional ByVal enabled As Boolean = True)
Public Sub ClearSingleCSTOutput()
```

Normal batch CST processing still saves to the original path:

```text
Desktop\QoS Automation\<DirPath>\NCA\<operator>\CST\...
```

Single CST processing uses the UserForm-selected output folder and `cstExcelName`.

The helper that decides the save path is:

```vba
Private Function CSTOutputWorkbookPath(...)
```

## CST Batch Processing

`CSTFiles.bas` remains the batch processor.

Current behavior:

- Reads campaign/trend from `Sheets("Option").Range("AO1").Value`.
- Scans:

```text
Desktop\QoS Automation\<Trends>\Telcos\<operator>\CST\LOGS
```

- Operators:

```text
MTN
TELECEL
AT
```

- Groups logs by extracted town/date.
- Splits files into MOC/MTC by reading `#DL`.
- Opens:

```text
Desktop\QoS Automation\Templates\Voice Template.xlsm
```

- Runs `CSTCOVMOS` for each file.

## Data Module Fixes

Files:

```text
DataFiles.bas
ThreeGDataMacro.bas
FourGDataMacro.bas
```

Key fixes made:

- Hardened data processing loops and cleanup paths.
- Added safer worksheet/file helper behavior.
- Reduced repeated pass loops to a single pass where appropriate.
- Stopped reopening workbook paths during cleanup.
- Added safer last-used-row handling.
- Added `KeepValue` helper changes so arrays returned by `Split(...)` can be passed safely.
- Fixed compile error:

```text
Type mismatch: array or user-defined type expected
```

Cause:

```vba
Keep = Split(...)
```

returns a Variant array, but the helper expected `String()`.

Fix:

```vba
Private Function KeepValue(ByVal Keep As Variant, ByVal index As Long) As String
```

- Fixed runtime error:

```text
Run-time error '424': Object required
```

Cause in `DataFiles.bas`:

```vba
If Not wb Is Nothing Then
```

`wb` was not declared as a `Workbook` in `ProcessDataFiles`.

Fix:

```vba
Dim wb As Workbook
```

## Coverage Module Fixes

Files:

```text
CovFiles.bas
ThreeGCovUI.bas
FourGCovAUI.bas
```

Coverage modules were reviewed and hardened earlier. Fixes were committed and pushed. Do not assume workbook modules contain the same code unless the user has imported the exported files manually.

Important note:

At one point, modules were mistakenly imported into workbooks for compile validation. The user later clarified:

```text
Next time please don't update in the workbook.
```

Respect this as the standing rule.

## Git Commit Trail

Recent commits:

```text
f678625 Clarify direct CST form helper call
c34311a Add direct CST form submit support
e93fd97 Rename CST COV MOS single macro module
1d07d89 Add single CST log processing flow
60381aa Fix DataFiles workbook cleanup object
ce329c4 Fix data KeepValue helper
d90a23d Harden data VBA processing
bff789a Add data processing modules
273ee74 Harden coverage VBA processing
2bac14d Add coverage UI modules
807a2c5 Initial VBA coding modules
```

## Validation Performed

Text-level checks were run on edited `.bas` files:

- Procedure and `End Sub` / `End Function` counts were balanced.
- Git status was checked after pushes.
- No workbook import/compile validation was performed after the user instructed not to update workbooks.

## Current Caveats

- The active IDE tab may still show `CSTCOVMOSMacro.bas`, but the exported source in the repo is now `CSTCOVMOSSingleMacro.bas`.
- Excel workbooks will not automatically receive these `.bas` changes. The user must manually import modules, or explicitly authorize workbook updates.
- If `ProcessSingleCSTFromSubmittedForm` is called from the UserForm, use:

```vba
ProcessSingleCSTFromSubmittedForm Me
```

- If the new UserForm is launched directly in the VBE without a caller, `Me.Hide` alone does not start processing. Either call the helper with `Me` in `lblSend_Click`, or launch the form using `OpenSingleCSTForm`.

