@echo off
setlocal enableextensions

REM Chrome DevTools launcher (WSL friendly)
set "PORT=%DEVTOOLS_PORT%"
if "%PORT%"=="" set "PORT=9222"
set "PROFILE_DIR=%USERPROFILE%\chrome-debug"

REM Ensure previous Chrome/Edge instances are not blocking the port
taskkill /IM chrome.exe /F >nul 2>nul
taskkill /IM msedge.exe /F >nul 2>nul

REM Locate Chrome installation
call :find_chrome_exe || goto :eof

echo [INFO] Chrome executable: "%CHROME_EXE%"
echo [INFO] Profile directory: "%PROFILE_DIR%"

REM Launch Chrome with remote debugging enabled
start "" "%CHROME_EXE%" --remote-debugging-port=%PORT% --remote-debugging-address=0.0.0.0 --user-data-dir="%PROFILE_DIR%"

echo [INFO] Chrome DevTools listening on 0.0.0.0:%PORT%
echo [INFO] Windows URL : http://127.0.0.1:%PORT%
echo [INFO] WSL URL     : http://%COMPUTERNAME%:%PORT% (or use resolv.conf detection)
echo [INFO] Update PLAYWRIGHT_CONNECT_URL if you need a fixed address.
exit /b 0

:find_chrome_exe
set "CHROME_EXE="
for %%P in (
    "C:\Program Files\Google\Chrome\Application\chrome.exe"
    "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%~P (
        set "CHROME_EXE=%%~P"
        goto :found_chrome
    )
)
echo [ERROR] Unable to find chrome.exe in standard locations.
echo         Please update start_chrome_dev.bat with your custom path.
pause
exit /b 1

:found_chrome
exit /b 0
