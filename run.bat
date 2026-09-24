@echo off
cd /d "%~dp0"

rem Uu tien lenh "py" (Python install manager), neu khong co thi dung "python"
set PY=py
where py >nul 2>nul || set PY=python
where %PY% >nul 2>nul || (
    echo Khong tim thay Python. Hay cai Python tu https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Dang cai thu vien (lan dau mat khoang 1 phut)...
%PY% -m pip install -q -r requirements.txt || (
    echo.
    echo Cai thu vien that bai. Kiem tra ket noi mang roi chay lai.
    pause
    exit /b 1
)

%PY% main.py || pause
