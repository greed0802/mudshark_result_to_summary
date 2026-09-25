@echo off
rem ================================================================
rem  SETUP_Python.bat  -  ONE-TIME setup
rem
rem  Downloads a PORTABLE Python (just a ZIP, no installer) and
rem  unpacks it into a "python" folder next to this file.
rem
rem     * NO admin rights needed
rem     * NO Microsoft Store needed
rem     * NO installer is run - it's a plain ZIP extraction
rem
rem  After this finishes once, use Run_Mudshark_Summary.bat forever.
rem ================================================================
setlocal EnableExtensions
title Portable Python setup (one-time)

set "DEST=%~dp0python"

if exist "%DEST%\python.exe" (
    echo Portable Python is already set up here:
    echo   %DEST%
    goto :verify
)

rem --- pick the right CPU build -------------------------------
set "PYARCH=amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "PYARCH=arm64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86"   set "PYARCH=win32"

echo.
echo  This will download a portable Python ZIP (~11 MB) from python.org
echo  and extract it to:
echo    %DEST%
echo  No admin rights, no Microsoft Store, no installer - just a ZIP.
echo.

set "ZIPTMP=%TEMP%\mudshark_python_embed.zip"
for %%V in (3.12.10 3.12.9 3.11.9) do (
    if not exist "%DEST%\python.exe" call :tryver %%V
)

if not exist "%DEST%\python.exe" (
    echo.
    echo ***************************************************************
    echo  The automatic download did not work ^(company network may be
    echo  blocking it^). You can do it MANUALLY instead - it is easy:
    echo.
    echo   1. In your web browser open:
    echo      https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip
    echo   2. Save the ZIP, right-click it, choose "Extract All..."
    echo   3. Extract it so that "python.exe" ends up inside a folder
    echo      named   python   next to this BAT file, i.e.:
    echo          %DEST%\python.exe
    echo   4. Run Run_Mudshark_Summary.bat as usual.
    echo ***************************************************************
    goto :end
)

:verify
echo.
echo Checking that it works:
"%DEST%\python.exe" --version
echo.
echo Setup complete. From now on just use Run_Mudshark_Summary.bat.
goto :end

rem ----------------------------------------------------------------
:tryver
set "URL=https://www.python.org/ftp/python/%~1/python-%~1-embed-%PYARCH%.zip"
echo Trying to download Python %~1 for %PYARCH% ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; [System.Net.WebRequest]::DefaultWebProxy.Credentials=[System.Net.CredentialCache]::DefaultCredentials; $ProgressPreference='SilentlyContinue'; try { Invoke-WebRequest -Uri '%URL%' -OutFile '%ZIPTMP%' -UseBasicParsing -ErrorAction Stop } catch { exit 1 }"
if errorlevel 1 (
    echo    download failed, will try another version...
    exit /b 0
)
echo Extracting ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Expand-Archive -LiteralPath '%ZIPTMP%' -DestinationPath '%DEST%' -Force -ErrorAction Stop } catch { exit 1 }"
if errorlevel 1 (
    echo    extraction failed, will try another version...
    rmdir /s /q "%DEST%" >nul 2>nul
    exit /b 0
)
del /q "%ZIPTMP%" >nul 2>nul
exit /b 0

:end
echo.
pause
