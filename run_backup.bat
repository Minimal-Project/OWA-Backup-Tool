@echo off
:: -------------------------------------------------------------------
:: 1) (Optional) Add Python 3.10 installation directory to your PATH
::    Remove or adjust this line if Python is already accessible on your PATH
:: -------------------------------------------------------------------
set "PATH=%LOCALAPPDATA%\Programs\Python\Python310\;%PATH%"

:: -------------------------------------------------------------------
:: 2) Change working directory to the location of this batch file
::    The %~dp0 variable expands to the drive letter and path of this script
:: -------------------------------------------------------------------
cd /d "%~dp0"

:: -------------------------------------------------------------------
:: 3) Ensure a "Downloads" folder exists in this directory
:: -------------------------------------------------------------------
set "OUTPUT_DIR=%~dp0Downloads"
if not exist "%OUTPUT_DIR%" (
    md "%OUTPUT_DIR%"
)

:: -------------------------------------------------------------------
:: 4) Launch the OWA downloader script, directing output to the Downloads folder
:: -------------------------------------------------------------------
python owa_downloader.py --output-dir "%OUTPUT_DIR%"

:: -------------------------------------------------------------------
:: 5) Pause execution so you can read any messages or errors before the window closes
:: -------------------------------------------------------------------
pause
