@echo off
rem ================================================================
rem  Run_Mudshark_Summary.bat
rem
rem  Easiest way to use on Windows:
rem     drag & drop the Mudshark Excel export (.xlsx) onto this file
rem
rem  or double-click it and paste the file path when asked.
rem ================================================================
setlocal EnableExtensions
title Mudshark Summary Builder

set "SCRIPT=%~dp0mudshark_summary.py"

if not exist "%SCRIPT%" (
    echo ERROR: mudshark_summary.py was not found next to this BAT file.
    echo Keep both files in the same folder.
    goto :end
)

if "%~1"=="" goto :ask
set "INPUT=%~1"
goto :have_input

:ask
echo.
echo  Drag the Mudshark export .xlsx onto this BAT file to use it,
echo  or paste the full path of the .xlsx below and press Enter.
echo.
set /p "INPUT=File path: "
rem remove surrounding quotes if the user pasted a quoted path
set "INPUT=%INPUT:"=%"

:have_input
if not exist "%INPUT%" (
    echo.
    echo ERROR: file not found:
    echo   %INPUT%
    goto :end
)

echo.
echo  Input file:
echo  %INPUT%
echo.

rem --- find a Python interpreter -------------------------------
where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%SCRIPT%" "%INPUT%"
    goto :done
)
where python >nul 2>nul
if not errorlevel 1 (
    python "%SCRIPT%" "%INPUT%"
    goto :done
)
where python3 >nul 2>nul
if not errorlevel 1 (
    python3 "%SCRIPT%" "%INPUT%"
    goto :done
)

echo ERROR: could not find Python on this computer.
echo.
echo You do NOT need admin rights to install it. Do ONE of these:
echo   1. Open the Microsoft Store app, search "Python 3.12", click Get.
echo   2. Or download python.org's Windows installer and tick
echo      "Install launcher for all users" OFF and install "for me only".
echo Then run this BAT file again.
goto :end

:done
echo.
echo The SUMMARY file is saved next to the Mudshark export.

:end
echo.
pause
