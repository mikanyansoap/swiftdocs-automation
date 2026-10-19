@echo off
setlocal

echo =====================================================================
echo  Running Pag-IBIG Forms in Development Mode (Briefcase Dev)
echo =====================================================================
echo.

if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

python -m briefcase dev

if %errorlevel% neq 0 (
    echo.
    echo Development launch encountered an issue. Running directly via python module...
    python -m pagibig_forms
)

pause
