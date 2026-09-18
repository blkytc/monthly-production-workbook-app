@echo off
setlocal
cd /d "%~dp0\.."
python -m PyInstaller --noconfirm --clean monthly-workbook.spec
if errorlevel 1 exit /b %errorlevel%
echo Built dist\MonthlyWorkbookGenerator.exe
