@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
if not exist "ket_qua_CSDL" mkdir "ket_qua_CSDL"
REM Keo tha file phieu phu luc da dien vao day. May kiem tra truoc, loi CHAN thi khong ghi.
if "%~1"=="" (
  echo Keo tha file phieu phu luc vao file .bat nay.
  pause
  exit /b
)
python csdl_hop_dong.py nap-phu-luc %*
pause
