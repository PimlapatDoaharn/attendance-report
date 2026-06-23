from __future__ import annotations

import PyInstaller.__main__

PyInstaller.__main__.run(
    [
        "run_app.py",
        "--name=AttendanceReportGenerator",
        "--onefile",
        "--windowed",
        "--paths=src",
        "--noconfirm",
    ]
)
