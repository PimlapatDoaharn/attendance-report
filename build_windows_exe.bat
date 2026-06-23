@echo off
setlocal

cd /d "%~dp0"

echo ========================================
echo Attendance Report Generator EXE Builder
echo ========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    set "PYTHON_CMD=python"
)

echo Creating Python virtual environment...
%PYTHON_CMD% -m venv .venv
if errorlevel 1 goto error

echo Installing required packages...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto error
pip install -r requirements.txt
if errorlevel 1 goto error

echo Building Windows executable...
python build_exe.py
if errorlevel 1 goto error

echo.
echo ========================================
echo Build completed successfully.
echo Send this file to HR:
echo dist\AttendanceReportGenerator.exe
echo ========================================
echo.
pause
exit /b 0

:error
echo.
echo Build failed. Please copy the error above and ask for help.
echo.
pause
exit /b 1