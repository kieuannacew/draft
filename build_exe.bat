@echo off
rem Đóng gói thành file .exe (không cần cài Python trên máy khác)
cd /d "%~dp0"
python -m pip install -q -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --windowed --name LuyenThiMOS main.py
echo.
echo Xong! File chay: dist\LuyenThiMOS.exe
pause
