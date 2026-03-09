@echo off
setlocal enabledelayedexpansion

:: ============================================
:: Hebrew Fonts Ribbon - Installer
:: ============================================

title Hebrew Fonts Ribbon Installer

color 0A
cls
echo.
echo    ╔════════════════════════════════════════╗
echo    ║  Hebrew Fonts Ribbon - Installer      ║
echo    ║  Version 1.0                           ║
echo    ╚════════════════════════════════════════╝
echo.

:: Configuration
set FILE_NAME=HebrewFontsRibbon.dotm
set DOWNLOAD_URL=https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%%20fonts.dotm
set STARTUP_FOLDER=%APPDATA%\Microsoft\Word\STARTUP
set TEMP_FILE=%TEMP%\%FILE_NAME%

:: Check if Word is running
echo [*] Checking if Microsoft Word is running...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    color 0C
    echo.
    echo  ╔════════════════════════════════════════╗
    echo ║  WARNING: Word is currently running!   ║
    echo ╚════════════════════════════════════════╝
    echo.
    echo Please close ALL Microsoft Word windows before installing.
    echo.
    choice /C YN /M "Close Word and continue installation"
    if errorlevel 2 (
        echo Installation cancelled.
        pause
        exit /b 0
    )
    color 0A
)

:: Check for existing installation
echo [*] Checking for existing installation...
if exist "%STARTUP_FOLDER%\%FILE_NAME%" (
    echo.
    echo Found existing installation.
    choice /C YN /M "Overwrite existing installation"
    if errorlevel 2 (
        echo Installation cancelled.
        pause
        exit /b 0
    )
)

:: Create STARTUP folder if needed
echo [*] Preparing installation folder...
if not exist "%STARTUP_FOLDER%" (
    echo     Creating STARTUP folder...
    mkdir "%STARTUP_FOLDER%" 2>NUL
    if errorlevel 1 (
        color 0C
        echo.
        echo ERROR: Cannot create STARTUP folder!
        echo Please run as Administrator.
        pause
        exit /b 1
    )
)

:: Download file
echo [*] Downloading from GitHub...
echo     Please wait...
echo.

powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%TEMP_FILE%' -UseBasicParsing } catch { exit 1 } }"

if errorlevel 1 (
    color 0C
    echo.
    echo ╔════════════════════════════════════════╗
    echo ║  ERROR: Download Failed!               ║
    echo ╚════════════════════════════════════════╝
    echo.
    echo Possible causes:
    echo  - No internet connection
    echo  - Invalid GitHub URL
    echo  - File not found on GitHub
    echo.
    echo Please check and try again.
    pause
    exit /b 1
)

if not exist "%TEMP_FILE%" (
    color 0C
    echo ERROR: File not downloaded!
    pause
    exit /b 1
)

:: Install file
echo [*] Installing to Word STARTUP folder...
copy /Y "%TEMP_FILE%" "%STARTUP_FOLDER%\%FILE_NAME%" >NUL

if errorlevel 1 (
    color 0C
    echo.
    echo ERROR: Installation failed!
    echo Cannot copy file to STARTUP folder.
    pause
    exit /b 1
)

:: Cleanup
echo [*] Cleaning up temporary files...
del "%TEMP_FILE%" 2>NUL

:: Success
cls
color 0A
echo.
echo    ╔════════════════════════════════════════╗
echo    ║  ✓ Installation Successful!            ║
echo    ╚════════════════════════════════════════╝
echo.
echo The Hebrew Fonts Ribbon has been installed to:
echo %STARTUP_FOLDER%
echo.
echo ╔════════════════════════════════════════╗
echo ║  NEXT STEPS:                           ║
echo ╚════════════════════════════════════════╝
echo.
echo  1. Close ALL Microsoft Word windows
echo  2. Restart Microsoft Word
echo  3. Look for the "Hebrew Fonts" tab in the ribbon
echo.
echo ╔════════════════════════════════════════╗
echo ║  ENABLE MACROS (First time only):     ║
echo ╚════════════════════════════════════════╝
echo.
echo  1. File → Options → Trust Center
echo  2. Trust Center Settings → Macro Settings
echo  3. Select "Enable all macros"
echo.
pause

:: Option to open STARTUP folder
echo.
choice /C YN /M "Open STARTUP folder to verify"
if not errorlevel 2 (
    explorer "%STARTUP_FOLDER%"
)

echo.
echo Thank you for installing Hebrew Fonts Ribbon!
pause
exit /b 0