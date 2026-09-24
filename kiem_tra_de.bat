@echo off
rem Keo tha thu muc de vao file nay de kiem tra de
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
if "%~1"=="" (
    echo Cach dung: keo tha thu muc de ^(trong de_thi^) vao file kiem_tra_de.bat
    echo.
    %PY% kiem_tra_de.py --luat
) else (
    %PY% kiem_tra_de.py "%~1"
)
pause
