@echo off
rem Đóng gói thành file .exe (không cần cài Python trên máy khác)
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -m pip install -q -r requirements.txt pyinstaller
%PY% -m PyInstaller --noconfirm --onefile --windowed --name LuyenThiMOS main.py
echo.
echo Xong! File chay: dist\LuyenThiMOS.exe
pause
