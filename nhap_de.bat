@echo off
rem Keo tha thu muc bo de (hoac file .zip) vao file nay de nhap de vao app
cd /d "%~dp0"
set PY=py
where py >nul 2>nul || set PY=python
if "%~1"=="" (
    echo Cach dung: keo tha thu muc bo de ^(vd MOS_Word365_DeThucTe^) hoac file .zip vao file nhap_de.bat
) else (
    %PY% nhap_de.py "%~1"
)
pause
