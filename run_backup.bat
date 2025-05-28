@echo off
setlocal enabledelayedexpansion

echo.
echo ===================================================================
echo                      ReHoGa Interactive
echo ===================================================================
echo                       OWA BACKUP TOOL
echo ===================================================================

:: -------------------------------------------------------------------
:: 1) (Optional) Add Python installation directory to your PATH
::     Remove or adjust this line if Python is already accessible on your PATH
:: -------------------------------------------------------------------
:: set "PATH=%LOCALAPPDATA%\Programs\Python\Python310\;%PATH%"

:: -------------------------------------------------------------------
:: 2) Change working directory to the location of this batch file
:: -------------------------------------------------------------------
cd /d "%~dp0"

:: -------------------------------------------------------------------
:: 3) Ensure a "Downloads" folder exists in this directory
:: -------------------------------------------------------------------
set "OUTPUT_DIR=%~dp0Downloads"
if not exist "%OUTPUT_DIR%" (
    echo Creating Downloads folder...
    md "%OUTPUT_DIR%"
)

:: -------------------------------------------------------------------
:: 4) Prompt user for mailbox type
:: -------------------------------------------------------------------
echo.
echo Choose mailbox type:
echo   1) Own mailbox
echo   2) Shared mailbox
set /p MAILBOX_CHOICE="Enter your choice (1 or 2): "

if "%MAILBOX_CHOICE%"=="2" (
    set "MAILBOX_TYPE=shared"
) else (
    set "MAILBOX_TYPE=own"
)

:: -------------------------------------------------------------------
:: 5) Launch the OWA downloader script with selected mailbox type
:: -------------------------------------------------------------------
echo.
echo Starting OWA Backup Tool with mailbox: %MAILBOX_TYPE%...
echo.

python owa_downloader.py --output-dir "%OUTPUT_DIR%" --mailbox %MAILBOX_TYPE%
set EXIT_CODE=%ERRORLEVEL%

if %EXIT_CODE% neq 0 (
    echo.
    echo The program encountered an error (Exit code: %EXIT_CODE%).
    echo Please check the messages above for more information.
) else (
    echo.
    echo Backup process completed.
    echo Emails have been saved to: %OUTPUT_DIR%
)
PAUSE
:: -------------------------------------------------------------------
:: 6) Pause execution before closing
:: -------------------------------------------------------------------
echo.
echo Press any key to exit...
pause >nul
