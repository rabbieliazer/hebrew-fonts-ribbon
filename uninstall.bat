@echo off
echo ============================================
echo Hebrew Fonts Ribbon - Uninstaller
echo ============================================
echo.

set FILE_NAME=HebrewFontsRibbon.dotm
set STARTUP_FOLDER=%APPDATA%\Microsoft\Word\STARTUP

if not exist "%STARTUP_FOLDER%\%FILE_NAME%" (
    echo Hebrew Fonts Ribbon is not installed.
    pause
    exit /b 0
)

echo Found installation at:
echo %STARTUP_FOLDER%\%FILE_NAME%
echo.

choice /C YN /M "Are you sure you want to uninstall"
if errorlevel 2 (
    echo Uninstall cancelled.
    pause
    exit /b 0
)

echo.
echo Removing Hebrew Fonts Ribbon...
del "%STARTUP_FOLDER%\%FILE_NAME%"

if errorlevel 1 (
    echo ERROR: Could not remove file!
    echo Please close Word and try again.
    pause
    exit /b 1
)

echo.
echo ✓ Uninstall complete!
echo.
echo Please restart Microsoft Word for changes to take effect.
pause