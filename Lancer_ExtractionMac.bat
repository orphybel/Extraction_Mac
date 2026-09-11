@echo off
REM Lance Extraction MAC. Necessite Python 3.8 ou plus recent (avec Tkinter).
cd /d "%~dp0"

where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -3 -m extraction_mac %*
) else (
    python -m extraction_mac %*
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Le programme ne s est pas lance correctement.
    echo Verifier que Python 3 est installe : https://www.python.org/downloads/windows/
    echo Lors de l installation, cocher "Add python.exe to PATH" et "tcl/tk and IDLE".
    pause
)
