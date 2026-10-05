@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
REM Module 2 - Lop 1: kiem tra tinh toan ven file ho so thanh toan (moi mau Excel)
REM Keo tha 1 hay nhieu file .xlsx vao day. Ket qua: ket_qua_toan_ven\toan_ven_<ten file>.xlsx
if "%~1"=="" (
  echo Keo tha file ho so thanh toan .xlsx vao file .bat nay.
  pause
  exit /b
)
if not exist "ket_qua_toan_ven" mkdir "ket_qua_toan_ven"
:lap
if "%~1"=="" goto xong
python toan_ven.py "%~1" --out "ket_qua_toan_ven\toan_ven_%~n1.xlsx"
start "" "ket_qua_toan_ven\toan_ven_%~n1.xlsx"
shift
goto lap
:xong
pause
