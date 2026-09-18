@echo off
setlocal
cd /d "%~dp0\.."
py -m PyInstaller --noconfirm --clean monthly-workbook.spec
if errorlevel 1 exit /b %errorlevel%
echo 已生成 dist\月度生产表生成器.exe
