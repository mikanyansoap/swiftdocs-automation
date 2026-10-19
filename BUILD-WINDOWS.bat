@echo off
setlocal

echo =====================================================================
echo  Building Windows Desktop Installer (.msi)
echo =====================================================================
echo.

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

:: Briefcase cannot use a tools folder whose path has spaces (for example
:: C:\Users\First Middle Last\...), so keep its tools in C:\BeeWare instead.
if not defined BRIEFCASE_HOME set "BRIEFCASE_HOME=C:\BeeWare\briefcase"
if not exist "%BRIEFCASE_HOME%" mkdir "%BRIEFCASE_HOME%"
echo Briefcase tools folder: %BRIEFCASE_HOME%
echo.

echo Creating Windows application scaffold...
python -m briefcase create windows --no-input
if errorlevel 1 goto :failed

echo.
echo Building Windows application...
python -m briefcase build windows --no-input
if errorlevel 1 goto :failed

echo.
echo Packaging Windows MSI installer...
python -m briefcase package windows --adhoc-sign --no-input
if errorlevel 1 goto :failed

echo.
echo =====================================================================
echo  Build finished!
echo  Your installer is in:
echo    %~dp0dist
echo =====================================================================
if exist "%~dp0dist" explorer "%~dp0dist"
pause
exit /b 0

:failed
echo.
echo =====================================================================
echo  BUILD FAILED - no installer was made.
echo  Scroll up to the first red error message to see why.
echo =====================================================================
pause
exit /b 1
