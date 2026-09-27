@echo off
setlocal
title Simplex Invoices and Quotes
pushd "%~dp0"
set "LAUNCHER=%~dp0thermofisher_invoice_app_simplex\launcher.py"

if exist "%LAUNCHER%" goto find_python
echo.
echo This launcher needs the rest of the app next to it.
echo If you opened it from inside the ZIP file: close this window, right click
echo the ZIP file, choose "Extract All...", then double click this file in the
echo extracted folder.
goto pause_fail

:find_python
set "PY="
call :try py -3.12
if not defined PY call :try py -3.13
if not defined PY call :try py -3.11
if not defined PY call :try py -3.14
if not defined PY call :try py -3.10
if not defined PY call :try python
if not defined PY call :try "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY call :try "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY call :try "%ProgramFiles%\Python313\python.exe"
if not defined PY call :try "%ProgramFiles%\Python312\python.exe"
if not defined PY goto no_python

%PY% "%LAUNCHER%" %*
if %errorlevel% NEQ 0 goto pause_fail
popd
exit /b 0

:try
%* -c "import platform, sys, sysconfig; ok = (3, 10) <= sys.version_info[:2] <= (3, 14) and platform.python_implementation() == 'CPython' and sysconfig.get_platform() == 'win-amd64' and not sysconfig.get_config_var('Py_GIL_DISABLED'); sys.exit(0 if ok else 1)" <nul >nul 2>&1
if %errorlevel% EQU 0 set PY=%*
exit /b 0

:no_python
echo.
echo 64 bit Python 3.10 to 3.14 was not found on this computer.
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" goto python_unusable
echo.
echo Python 3.13 can be installed now for your account only (no administrator
echo rights needed) with the winget tool that comes with Windows.
choice /C YN /M "Install Python 3.13 now"
if errorlevel 2 goto manual_python
winget install -e --id Python.Python.3.13 --scope user --architecture x64 --accept-package-agreements --accept-source-agreements
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" goto find_python

:manual_python
echo.
echo Install Python 3.13 from https://www.python.org/downloads/windows/
echo (the page opens now): choose "Windows installer (64-bit)", run it, then
echo double click this file again.
start "" "https://www.python.org/downloads/windows/"
goto pause_fail

:python_unusable
echo Python 3.13 is installed in "%LOCALAPPDATA%\Programs\Python\Python313"
echo but does not start. Repair or reinstall it from
echo https://www.python.org/downloads/windows/ and then double click this file again.
goto pause_fail

:pause_fail
echo.
pause
popd
exit /b 1
