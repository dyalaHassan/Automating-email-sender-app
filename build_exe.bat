@echo off
REM =========================================================
REM build_exe.bat
REM
REM Builds a single standalone .exe of the Bulk Email Sender.
REM Run this ONCE, on a Windows PC with Python installed, from
REM inside the project folder (the one containing app.py).
REM
REM The result will be in:  dist\OneWorldBulkEmailSender.exe
REM That one file is what you hand out to employees - they do
REM NOT need Python, Flask, or anything else installed.
REM
REM This version calls everything through "python -m ..." so it
REM works even if pyinstaller.exe isn't on your system PATH.
REM =========================================================

echo Installing/upgrading required packages...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo.
    echo ERROR: pip itself failed to run. Is Python installed and
    echo added to PATH? Try running "python --version" to check.
    pause
    exit /b 1
)

python -m pip install --upgrade pyinstaller
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install PyInstaller. See the message above.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install requirements.txt packages.
    pause
    exit /b 1
)

echo.
echo Building the .exe (this can take a few minutes)...
python -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --name OneWorldBulkEmailSender ^
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

REM ---------------------------------------------------------
REM Only claim success if the .exe actually exists.
REM ---------------------------------------------------------
if exist "dist\OneWorldBulkEmailSender.exe" (
    echo.
    echo SUCCESS. Your .exe is at: dist\OneWorldBulkEmailSender.exe
    echo Copy that single file anywhere and double-click it to run.
) else (
    echo.
    echo ERROR: Build finished but dist\OneWorldBulkEmailSender.exe
    echo was not found. Scroll up for the real error message.
)

pause
