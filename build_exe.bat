@echo off
rem Đóng gói thành file .exe (không cần cài Python trên máy khác)
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
%PY% -m pip install -q -r requirements.txt pyinstaller
%PY% -m PyInstaller --noconfirm --onefile --windowed --name LuyenThiMOS --add-data "mos\i18n_en.json;mos" --add-data "mos\dich_de.json;mos" --add-data "mos\qt\fonts;mos\qt\fonts" main.py
rem Chep kem thu muc de thi (de mau + bo de da nhap) canh file .exe
xcopy /e /i /y /q de_thi dist\de_thi >nul
xcopy /e /i /y /q tai_lieu dist\tai_lieu >nul
echo.
echo Xong! File chay: dist\LuyenThiMOS.exe (chep ca thu muc dist\de_thi va dist\tai_lieu di kem)
pause
