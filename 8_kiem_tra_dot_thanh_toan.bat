@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
if not exist "ket_qua_CSDL" mkdir "ket_qua_CSDL"
REM Keo tha file ho so thanh toan dot moi vao day. Doi chieu voi BOQ + phu luc hieu luc tai ngay nghiem thu.
if "%~1"=="" (
  echo Keo tha file ho so thanh toan .xlsx vao file .bat nay.
  pause
  exit /b
)
set /p HD=Nhap so hop dong (vd 0404):
set /p NT=Ngay nghiem thu cua dot nay (dd/mm/yyyy):
set /p DOT=So dot (Enter de may tu doan tu ten file):
if "%DOT%"=="" (
  python csdl_hop_dong.py kiem-dot "%HD%" "%~1" --ngay-nghiem-thu %NT% --out "ket_qua_CSDL\kiem_dot_%~n1.xlsx"
) else (
  python csdl_hop_dong.py kiem-dot "%HD%" "%~1" --ngay-nghiem-thu %NT% --dot %DOT% --out "ket_qua_CSDL\kiem_dot_%~n1.xlsx"
)
if errorlevel 1 (pause & exit /b)
start "" "ket_qua_CSDL\kiem_dot_%~n1.xlsx"
pause
