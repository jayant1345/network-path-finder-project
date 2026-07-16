@echo off
setlocal enabledelayedexpansion
title CPAN Network Path Finder

:: Always run from the folder this .bat file lives in, no matter where the
:: project folder is copied/moved to (D:\, C:\, a USB stick, etc).
cd /d "%~dp0"

set "VENV_DIR=venv"
set "LOCK_FILE=%VENV_DIR%\requirements.lock"

:: -- Pick a Python launcher --------------------------------------------
set "PY_CMD="
where python >nul 2>&1 && set "PY_CMD=python"
if not defined PY_CMD (
    where py >nul 2>&1 && set "PY_CMD=py -3"
)
if not defined PY_CMD (
    echo ERROR: Python was not found on PATH.
    echo Install Python 3.10+ from https://www.python.org/downloads/ and make sure
    echo "Add python.exe to PATH" is checked during install, then run this file again.
    pause
    exit /b 1
)

:: -- Decide whether the environment needs to be rebuilt ------------------
set "NEEDS_SETUP=0"
if not exist "%VENV_DIR%\Scripts\python.exe" set "NEEDS_SETUP=1"
if not exist "%LOCK_FILE%" set "NEEDS_SETUP=1"
if exist "%LOCK_FILE%" (
    fc /b "requirements.txt" "%LOCK_FILE%" >nul 2>&1
    if errorlevel 1 set "NEEDS_SETUP=1"
)

if "%NEEDS_SETUP%"=="1" (
    echo ============================================================
    echo First run on this machine / requirements changed.
    echo Rebuilding the Python environment - this only happens once
    echo and may take a few minutes. Please wait...
    echo ============================================================

    if exist "%VENV_DIR%" (
        echo Removing old "venv" folder...
        attrib -r -h -s "%VENV_DIR%\*.*" /s /d >nul 2>&1
        for /l %%i in (1,1,5) do (
            if exist "%VENV_DIR%" (
                rmdir /s /q "%VENV_DIR%" 2>nul
                ping -n 2 127.0.0.1 >nul
            )
        )
        if exist "%VENV_DIR%" (
            echo ERROR: Could not fully delete the old "venv" folder - some files
            echo are still locked. Close any other Command Prompt, VS Code, or
            echo Python process that might be using it, including a previous
            echo copy of this server still running, then run this file again.
            pause
            exit /b 1
        )
    )
    :: clean up the legacy folder name used by older versions of this project
    if exist "venv1" (
        echo Removing legacy "venv1" folder from a previous version...
        attrib -r -h -s "venv1\*.*" /s /d >nul 2>&1
        rmdir /s /q "venv1" 2>nul
        if exist "venv1" (
            echo WARNING: Could not remove the old "venv1" folder - it is locked
            echo by another process. It will just sit there unused; you can
            echo delete it manually later.
        )
    )

    echo Creating virtual environment...
    %PY_CMD% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create the virtual environment.
        pause
        exit /b 1
    )

    call "%VENV_DIR%\Scripts\activate.bat"

    echo Installing dependencies from requirements.txt...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Dependency installation failed. Check your internet connection
        echo and re-run START_SERVER.bat.
        pause
        exit /b 1
    )

    echo Installing Playwright browser - needed for the Docket Report feature...
    python -m playwright install chromium
    if errorlevel 1 (
        echo WARNING: Playwright browser download failed. The Docket Report
        echo feature will not work until this succeeds. All other features
        echo will still run normally. Re-run this file later to retry.
    )

    copy /y "requirements.txt" "%LOCK_FILE%" >nul
    echo.
    echo Environment setup complete.
    echo.
) else (
    call "%VENV_DIR%\Scripts\activate.bat"
)

if not exist ".env" (
    echo ------------------------------------------------------------
    echo NOTE: No .env file found. The Docket Report feature needs one
    echo with GCS_USERNAME, GCS_PASSWORD and GCS_LOGIN_URL set.
    echo The Network Path Finder itself will still work fine.
    echo ------------------------------------------------------------
)

echo.
echo Server starting at http://127.0.0.1:5001
echo Press Ctrl+C to stop the server.
echo.

start "" http://127.0.0.1:5001
python server.py

pause
