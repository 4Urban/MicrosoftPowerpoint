@echo off
REM Get User Input
echo Type the Number which you want!
echo.
echo   1. Normal(:= 96 dpi)
echo.
echo   2. Max(:= 300 dpi)
echo.
echo.

set /p id="Entry: "
echo %id%
pause
REM Set the Resolution
IF exist "%id%"(
    echo Max
    set decentry=300
    set resolution="300 dpi%
) ELSE (
    echo Normal
    set decentry=96
    set resolution="96 dpi%"
)

REM Edit the Registry
REG ADD HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options /v ExportBitmapResolution /t REG_DWORD /d %decentry% /f

REM Finish the Program
echo The Powerpoint (to PNG) Resolution has Change to %resolution%
pause
