# Attendance Report Generator

Mini standalone-friendly Python program for generating an updated Attendance Report workbook from:

1. `Report Finger Scan (.xlsx)` raw scan export
2. `[Draft] Attendance Report (.xlsx)` template/report workbook
3. `PTVN Report finger scan 2026 (.xlsx/.xlsm)` department/individual summary workbook

The app does not use AI. It uses deterministic normalization and fuzzy text matching for employee names.

## Features

- Adds/rebuilds `Raw finger scan`
- Adds/rebuilds `Count scan`
- Updates latest month fields in `2026 Report`
- Parses mixed date/time formats, including common Thai month names
- Fuzzy-matches imperfect employee names against `ฐานข้อมูลพนักงาน (HR)`
- Resolves suspicious employee IDs using `Previous month` hints first, then most common ID in the scan file, then HR ID
- Counts only unique weekday scan dates, so multiple scans in one day count as `1`
- Rounds `Leave / Off-site work` totals up when fractions such as `0.5` exist
- Uses the `Working days` value entered in the UI as the denominator for `Target`
- Updates `PTVN Report finger scan 2026` summary and individual attendance sheets
- Dashboard-style UI with summary cards, progress status, and Open File/Open Folder actions
- Preserves `.xlsm` PTVN templates and lets VBA create dashboard charts without openpyxl chart XML

## Expected Workbook Sheets

The draft Attendance Report should include:

- `2026 Report`
- `Count scan` (will be rebuilt)
- `ข้อมูลวันลา`
- `ฐานข้อมูลพนักงาน (HR)`
- `Previous month`

The output workbook will contain:

- `2026 Report`
- `Count scan`
- `ข้อมูลวันลา`
- `ฐานข้อมูลพนักงาน (HR)`
- `Raw finger scan`
- `Previous month`

The PTVN workbook should include:

- `2026 Summary Report`
- `2026 Individual Attendance Repo`

The generated PTVN output will update:

- `2026 Summary Report`: department employee count, selected month working day, and average department attendance
- `2026 Individual Attendance Repo`: selected month working day, leave/off-site work, target 60%, and attendance percentage by employee

## Run From Source

```zsh
cd /Users/pimlapat.doa/Documents/Attendance_Report
/usr/local/bin/python3 -m pip install -r requirements.txt
PYTHONPATH=src /usr/local/bin/python3 run_app.py
```

## Build Standalone App / EXE

Install dependencies first, then run:

```zsh
cd /Users/pimlapat.doa/Documents/Attendance_Report
/usr/local/bin/python3 -m pip install -r requirements.txt
PYTHONPATH=src /usr/local/bin/python3 build_exe.py
```

Output will be under `dist/`.

- On macOS, PyInstaller creates a macOS executable.
- To create a Windows `.exe`, run the same project on Windows with Python installed.

## Notes

- The app now creates the output automatically next to the selected Attendance Report file, named like `Attendance_Report_May2026.xlsx`.
- The app also creates a PTVN output automatically next to the selected PTVN file, named like `PTVN_Report_Finger_Scan_May2026.xlsx` or `.xlsm` when the selected PTVN file is macro-enabled.
- `Working days` controls the target formula: `(Working Day + Leave / Off-site work) / Working days`.
- The same `Working days` value is written to row 3 for the selected month in both Attendance Report and PTVN workbooks.
- Click `Auto` to fill `Working days` from the number of Monday-Friday dates in the selected month.
- If matching is too strict, lower the fuzzy threshold in the UI, e.g. `75`.
- If wrong employees are matched, raise the threshold, e.g. `90`.
- The tool tries to detect common column names automatically, but consistent headers improve accuracy.

## Optional `.xlsm` Dashboard Chart Template

For the most stable dashboard charts, use a macro-enabled PTVN template:

1. Open the PTVN workbook in Excel and save it as `.xlsm`.
2. Press `Option + F11` / `Alt + F11` to open the VBA editor.
3. Import `vba/PTVN_DashboardCharts.bas` as a standard module.
4. Save the workbook as your PTVN `.xlsm` template.
5. Select that `.xlsm` file in the app.

When the PTVN input is `.xlsm`, Python writes the Dashboard formulas and helper ranges but does not create chart objects. The `RefreshPTVNDashboardCharts` VBA macro creates the charts using Excel's native chart engine when macros are enabled.
