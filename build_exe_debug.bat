@echo off
REM =========================================================
REM build_exe_debug.bat
REM
REM Same build as build_exe.bat, EXCEPT it keeps the console
REM window visible so any startup error/traceback is shown
REM instead of hidden. Use this ONLY to diagnose a problem.
REM Once things work, go back to using build_exe.bat for the
REM real version you hand out to employees.
REM =========================================================

echo Building a DEBUG .exe (console window will stay visible)...
python -m PyInstaller ^
    --onefile ^
    --name OneWorldBulkEmailSender_DEBUG ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import=win32timezone ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    app.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: PyInstaller failed. Scroll up to see the actual error.
    pause
    exit /b 1
)

if exist "dist\OneWorldBulkEmailSender_DEBUG.exe" (
    echo.
    echo SUCCESS. Now go run:
    echo   dist\OneWorldBulkEmailSender_DEBUG.exe
    echo A black console window will stay open - if it crashes,
    echo screenshot whatever text/traceback appears in it and send
    echo it back. The window will NOT auto-close on error so you
    echo have time to read/screenshot it.
) else (
    echo.
    echo ERROR: Build finished but the .exe was not found.
)

pause
