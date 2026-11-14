@echo off
cd /d "%~dp0"
echo ================================================
echo   Safety Helper - first-time setup (1 time only)
echo ================================================
echo.
set PY=
where py >nul 2>nul && set PY=py
if not defined PY (where python >nul 2>nul && set PY=python)
if not defined PY goto nopython
%PY% --version
if errorlevel 1 goto nopython
echo.
echo [Installing libraries... please wait 1-3 minutes]
%PY% -m pip install --upgrade pandas openpyxl pypdf python-docx olefile
if errorlevel 1 goto fail
echo.
echo ================================================
echo   DONE!  Now double-click  the "2_...(run).bat" file
echo ================================================
pause
exit /b

:nopython
echo.
echo [!] Python is not installed.
echo     1) Opening python.org - download and install Python 3
echo     2) On the first screen, CHECK "Add python.exe to PATH"
echo     3) Then double-click this file again.
start "" https://www.python.org/downloads/
pause
exit /b

:fail
echo.
echo [!] Install failed. Check internet connection.
echo     On a company network, the proxy/firewall may block it - ask IT.
pause
