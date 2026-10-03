@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
if not exist "ket_qua_CSDL" mkdir "ket_qua_CSDL"
REM Keo tha 1 hay nhieu file master_HD_xxx.xlsx vao day de nap vao co so du lieu
if "%~1"=="" (
  echo Keo tha file master_HD_xxx.xlsx vao file .bat nay.
  pause
  exit /b
)
python csdl_hop_dong.py nap-master %*
pause
