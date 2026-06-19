# VBA Coding Handoff

Last updated: 2026-06-19

Repository: https://github.com/Clive-B/VBA-Coding.git

Workspace:

```text
C:\Users\codro\Desktop\VBA Coding
```

## Operating Rules

- Edit the exported `.bas` / `.txt` files in this folder first.
- Do not import modules into live Excel workbooks unless the user explicitly asks.
- Before risky changes, create timestamped `.bak-YYYYMMDD-*` backups beside the source file.
- The user manually imports updated modules into the relevant workbook/template.
- Avoid public globals for cross-module state where possible; pass values as procedure arguments to prevent `Ambiguous name detected`.
- Git may warn about `C:\Users\codro/.config/git/ignore` permission. It has not blocked normal commits.

## UserForm Pattern

The current Data/Coverage forms use image-backed UI with transparent label hot zones.

Common control names:

```text
txtDataExcelName / txtCoverageExcelName
lblBrowseDataFile / lblBrowseCoverageFile
lblSelectFolder
lbl3G
lbl4G
lblSend
lbl3GLeft
lbl3GRight
lbl4GLeft
lbl4GRight
```

The network labels act as invisible buttons. Marker labels beside 3G/4G are used for hover/selected color feedback. When a network is clicked, that selection stays highlighted, and the other network loses its highlight.

For coverage form code, see:

```text
CoverageUserFormCode.txt
```

## CST Single Processing

Main files:

```text
CSTSingleFileProcessing.bas
CSTCOVMOSSingleMacro.bas
SaveWorkbookMacro.bas
ResultsWB.bas
```

Key updates:

- Added UserForm-driven single CST log processing.
- The selected CST log is used as an anchor to find related MOC/MTC files in the same folder.
- `CSTCOVMOSSingleMacro.bas` was separated from the batch `CSTCOVMOSMacro` module to avoid duplicate procedure/module conflicts.
- Single and batch CST can exist in the same workbook, but duplicate public names still need care.
- `SaveTelcoWorkbook` was changed to support optional output folder/campaign path:

```vba
Sub SaveTelcoWorkbook(Optional ByVal outputFolderPath As String = "", Optional ByVal dirPathValue As String = "")
```

- For single output, save to the UserForm-selected output folder.
- For batch output, preserve the original campaign structure.
- `DirPathGlobal` ambiguity was avoided by passing `dirPath` into procedures instead of declaring duplicate globals.
- `Copy_CSTMOS_TO_Results` is not part of `CSTCOVMOSSingleMacro`.
- Single CST false-fatal behavior was corrected by exiting before final-only work when the current log is not the final related file:

```vba
If Not Finish Then
    Exit Sub
End If
```

## 3G Data Single Processing

Main files:

```text
ThreeGDataSingleFileProcessing2.bas
ThreeGDataMacro2.bas
SaveWorkbookDataMacro.bas
```

Important behavior change:

- The selected file is now a day/technology selector, not a one-location selector.
- Example data filename:

```text
26Jun09 112358 BEACH-ROAD ABSA BANK DATA DAY 1.2.nmf
```

- The single coordinator now groups related files by:

```text
Date | DATA DAY
```

instead of:

```text
Date | Town | Location
```

- It still filters candidate logs by `#DL` technology, so only 3G files are processed.
- Each file still passes its own location to `ThreeGDataSingle`, allowing spot creation per location.
- Earlier files run with `Finish = False`; only the final selected-day file runs final save/post-processing.

Patch details:

- `ThreeGDataSingleFileProcessing2.bas`
  - Added `ExtractDayFromDataFile`.
  - Updated `BuildSingleDataKey` to use `ExtractDateFromDataFile & "|" & ExtractDayFromDataFile`.
- `ThreeGDataMacro2.bas`
  - Added/kept single-output support.
  - Changed location-boundary spot write so single mode always records the completed previous spot:

```vba
If SingleDataOutputEnabled Or Dir(...) = "" Then
```

  - Skipped deleting generated spot sheets during single mode:

```vba
If Not SingleDataOutputEnabled Then
    ' delete SpotOne, SpotTwo, ...
End If
```

Test finding:

- `TELECEL BEACH-ROAD 3G DATA DAY 1.xlsm` showed both logs were processed in `Logs`, but only one spot appeared in `Locations`.
- Cause was the batch-only `Dir(...) = ""` guard preventing the first completed location from being written.
- Fixed in `ThreeGDataMacro2.bas`.

## 4G Data Single Processing

Main files:

```text
FourGDataSingleFileProcessing.bas
FourGDataMacro.bas
SaveWorkbookFourGMacro.bas
```

Key updates:

- Added 4G single data processing route.
- `FourGDataSingleFileProcessing.bas` now mirrors the corrected 3G behavior.
- The selected file groups related 4G data logs by:

```text
Date | DATA DAY
```

- `ExtractDayFromDataFile` was added.
- `BuildSingleDataKey` was changed from date/town/location to date/day.
- Candidate logs are still filtered by `#DL = 4G`.
- `FourGDataMacro.bas` now records previous completed spots in single mode:

```vba
If Single4GDataOutputEnabled Or Dir(...) = "" Then
```

- 4G already had protection around deleting `Spot...Ping`, `Spot...Http`, and `Spot...FTP` sheets:

```vba
If Not Single4GDataOutputEnabled Then
```

so generated spot sheets remain in single mode.

## Data Save Macros

Main files:

```text
SaveWorkbookDataMacro.bas
SaveWorkbookFourGMacro.bas
SaveWorkbookMacro.bas
```

Key updates:

- Save procedures accept optional output folder and campaign path.
- Single mode saves copied workbooks into the UserForm-selected folder.
- Batch mode keeps the original campaign folder structure.
- `SaveWorkbookDataMacro.bas` maintains copied header formatting/font color by preserving the original copy/paste behavior for relevant ranges.
- For 3G data copied workbook, the header formatting from:

```vba
ThisWorkbook.Sheets("GSM_DataReport_multimetric_5").Range("K1:R1").Copy
NetworkWorkbook.Sheets(WSVame).Range("A1:H1").PasteSpecial
```

was preserved.

## Coverage Single Processing

Main files discussed or generated:

```text
ThreeGCovSingleFileProcessing.bas
ThreeGCovSingle.bas
SaveWorkbookThreeGCovMacro.bas
FourGCovSingleFileProcessing.bas
FourGCovSingle.bas
SaveWorkbookFourGCovMacro.bas
CoverageUserFormCode.txt
```

Notes:

- Some files were named by the user for convenience before full code existed.
- UserForm code/module code was generated for coverage single processing.
- Single coverage mode should skip frequency comparison procedures that depend on batch campaign state:
  - 3G: `GetARFCNValues`
  - 4G: EARFCN equivalent where applicable
- Single coverage mode needs fallback handling for empty `NetworkCurr`.
- AT does not currently have 4G coverage, so 4G coverage behavior was kept conservative.

## ARFCN / EARFCN Filtering

Files reviewed:

```text
CodeUI.bas
AutoFilterWithArrayUI.bas
ARFCNComp.bas
AutoFilterWithArrayUIFourG.bas
EARFCNComp.bas
```

Outcome:

- Batch-only ARFCN/EARFCN comparison steps should be skipped for single processed coverage files.
- Single mode lacks some batch variables like `NetworkCurr`; fallback values must be supplied or the dependent procedure skipped.

## Important Import Notes

For 3G data single fixes, import:

```text
ThreeGDataSingleFileProcessing2.bas
ThreeGDataMacro2.bas
```

The template workbook expects the callable module name used in `Application.Run`, for example:

```vba
ThreeGDataMacro.ConfigureSingleDataOutput
ThreeGDataMacro.ThreeGDataSingle
```

So after import, ensure the module name inside the template matches what the coordinator calls.

For 4G data single fixes, import:

```text
FourGDataSingleFileProcessing.bas
FourGDataMacro.bas
```

The 4G template expects:

```vba
FourGDataMacro.ConfigureSingle4GDataOutput
FourGDataMacro.FourGDataSingle
```

## Backups Created During This Work

Recent backup examples:

```text
ThreeGDataSingleFileProcessing2.bas.bak-20260618-daywide-single-data
ThreeGDataMacro2.bas.bak-20260618-daywide-single-data
ThreeGDataMacro2.bas.bak-20260619-single-spot-boundary
ThreeGDataMacro2.bas.bak-20260619-keep-single-spot-sheets
FourGDataSingleFileProcessing.bas.bak-20260619-daywide-single-data
FourGDataMacro.bas.bak-20260619-single-spot-boundary
```

Older backup files remain in the workspace for CST, data, coverage, save macros, ARFCN, and EARFCN work.

## Validation Performed

- Text-level procedure tracing was performed with PowerShell.
- Excel COM was used in read-only mode to inspect `TELECEL BEACH-ROAD 3G DATA DAY 1.xlsm`.
- That inspection confirmed both 3G logs were processed but only one spot was initially written.
- After analysis, the spot-boundary logic was patched.
- No direct workbook module import/compile was performed unless the user explicitly does it manually.

## Current Caveats

- Workbooks do not automatically receive `.bas` changes. Import updated modules manually into the correct workbook/template.
- If both original and single modules live in one workbook, avoid duplicate public procedure names unless `Application.Run` qualifies the exact module.
- Generated test workbooks should not be committed unless intentionally needed as artifacts.
- If single processing appears to process all logs but creates too few spots, check:
  - the related-file grouping key,
  - `prevlocKey` / `PrevlocationName`,
  - location-boundary spot writes,
  - and any old `Dir(...) = ""` batch guards.
