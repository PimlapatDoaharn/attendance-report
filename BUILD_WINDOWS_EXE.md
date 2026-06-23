# Build Windows EXE for HR

This project must be built on the same operating system that will run it.

- On macOS, PyInstaller creates a macOS app: `dist/AttendanceReportGenerator.app`.
- On Windows, PyInstaller creates a Windows executable: `dist\AttendanceReportGenerator.exe`.

## Recommended Build Steps

1. Copy the whole `Attendance_Report` project folder to a Windows computer.
2. Install Python 3.11+ from <https://www.python.org/downloads/windows/>.
3. Open the `Attendance_Report` folder on Windows.
4. Double-click `build_windows_exe.bat`.
5. Wait until it says `Build completed successfully`.

## File To Send To HR

Send this file:

```text
dist\AttendanceReportGenerator.exe
```

HR can double-click `AttendanceReportGenerator.exe` to open the app.

## If Windows Blocks The App

Because the file is not code-signed, Windows SmartScreen may show a warning.

Use:

```text
More info → Run anyway
```

For company-wide distribution, ask IT to code-sign the `.exe`.

## Build Command Manually

If the `.bat` file cannot be used, run these commands in Windows Command Prompt:

```bat
cd path\to\Attendance_Report
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python build_exe.py
```

The output will be:

```text
dist\AttendanceReportGenerator.exe
```