@echo off
echo Starting COA Add-in for IE11...

REM Kill existing Word processes
taskkill /F /IM WINWORD.EXE 2>nul

REM Set IE11 as default for Office
reg add "HKCU\Software\Microsoft\Office\16.0\Common\COM Compatibility\{8856F961-340A-11D0-A96B-00C04FD705A2}" /v "Compatibility Flags" /t REG_DWORD /d 0 /f

REM Clear Office cache
del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.*" 2>nul

REM Start the dev server
start cmd /k "npm run dev"

REM Wait for server to start
timeout /t 5

REM Start Word with the add-in
start winword

echo.
echo Add-in server started. Please load the add-in manually in Word:
echo 1. Go to Insert - My Add-ins
echo 2. Select COA Processor
echo.
pause