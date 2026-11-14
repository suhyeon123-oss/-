@echo off
cd /d "%~dp0"
where pyw >nul 2>nul && (start "" pyw "%~dp0safety_helper.pyw" & exit /b)
where pythonw >nul 2>nul && (start "" pythonw "%~dp0safety_helper.pyw" & exit /b)
echo [!] Python not found. Double-click the "1_...(setup).bat" file first.
pause
