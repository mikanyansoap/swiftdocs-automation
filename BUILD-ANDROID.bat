@echo off
setlocal

echo =====================================================================
echo  Building Android App (Debug .apk)
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

echo Creating Android scaffold...
python -m briefcase create android --no-input
if errorlevel 1 goto :failed

echo.
echo Building Android debug APK...
python -m briefcase build android --no-input
if errorlevel 1 goto :failed

echo.
echo Packaging Android APK...
python -m briefcase package android --no-input
if errorlevel 1 goto :failed

echo.
echo =====================================================================
echo  Android build finished!
echo  Your app package is in:
echo    %~dp0dist
echo =====================================================================
if exist "%~dp0dist" explorer "%~dp0dist"
pause
exit /b 0

:failed
echo.
echo =====================================================================
echo  BUILD FAILED - no app package was made.
echo  Scroll up to the first red error message to see why.
echo =====================================================================
pause
exit /b 1
