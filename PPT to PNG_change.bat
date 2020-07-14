@echo off
REM Get User Input
echo Type the Number which you want!
echo.
echo   1. 50 dpi
echo.
echo   2. 96 dpi = NORMAL
echo.
echo   3. 100 dpi
echo.
echo   4. 150 dpi
echo.
echo   5. 200 dpi
echo.
echo   6. 250 dpi
echo.
echo   7. 300 dpi = MAX
echo.
echo.

set /p resolution="Entry: "

REM Edit the Registry
REG ADD HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options /v ExportBitmapResolution /t REG_DWORD /d %resolution% /f

REM Finish the Program
echo Powerpoint (to PNG) Resolution has Change to %resolution% dpi.
pause
