@echo off
chcp 65001 >nul
REM ================================================================
REM  Trich thong tin MOT hop dong bang AI (may Local AI trong LAN) va so voi quy tac
REM  Cach dung: KEO THA file hop dong (.doc/.docx/.pdf) vao file nay
REM  Dia chi may AI, model: cau_hinh_may_ai.json (sua bang 14_cau_hinh_may_AI.bat)
REM  Nhieu hop dong cung luc: dung 15_trich_lo_hop_dong.bat
REM ================================================================
if "%~1"=="" (
  echo Keo tha file hop dong vao file .bat nay.
  pause
  exit /b
)
cd /d %~dp0
if not exist "ket_qua_AI" mkdir "ket_qua_AI"
python -m pip install -q openpyxl python-docx pdfplumber pywin32
python master_hd.py ai-trich "%~1" --an-danh --out "ket_qua_AI\so_sanh_AI_%~n1.xlsx"
echo.
echo Ket qua: ket_qua_AI\so_sanh_AI_%~n1.xlsx
pause
