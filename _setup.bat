@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo  Pag-IBIG Forms Setup - Summerfield Resources Corporation
echo =====================================================================
echo.

:: 1. Check for Python
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python not found. Installing Python 3.12 via Windows Package Manager (winget)...
    winget install --id Python.Python.3.12 -e --accept-source-agreements --accept-package-agreements
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install Python automatically.
        echo Please download and install Python 3.12 manually from https://python.org
        pause
        exit /b 1
    )
    echo Python installed. Refreshing PATH...
    call refreshenv >nul 2>nul
)

python --version
echo.

:: 2. Create Virtual Environment (.venv)
if not exist ".venv" (
    echo Creating virtual environment (.venv)...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Could not create virtual environment.
        pause
        exit /b 1
    )
    echo Virtual environment created.
) else (
    echo Virtual environment (.venv) already exists.
)

:: 3. Activate and Install Dependencies
echo.
echo Installing and updating core dependencies (Briefcase, Toga 0.5, PyPDF, OpenPyXL, Anthropic SDK, Pillow)...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install briefcase "toga>=0.5.7,<0.6" pypdf openpyxl "anthropic>=1.15.0,<2" Pillow

echo.
echo =====================================================================
echo  Setup Completed Successfully!
echo  You can now run:
echo    - TRY-ON-PC.bat      (to test the application on this computer)
echo    - BUILD-WINDOWS.bat  (to build the Windows installer .msi)
echo    - BUILD-ANDROID.bat  (to build the Android APK)
echo =====================================================================
echo.
pause
