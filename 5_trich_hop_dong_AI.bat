@echo off
chcp 65001 >nul
REM ================================================================
REM  Trich thong tin hop dong bang AI (Ollama qua tunnel) va so voi quy tac
REM  Cach dung: KEO THA file hop dong (.doc/.docx/.pdf) vao file nay
REM  Dieu kien: dang mo cua so 2_mo_giao_dien.bat (tunnel toi may AI)
REM ================================================================
if "%~1"=="" (
  echo Keo tha file hop dong vao file .bat nay.
  pause
  exit /b
)
cd /d %~dp0
if not exist "ket_qua_AI" mkdir "ket_qua_AI"
python -m pip install -q openpyxl python-docx pdfplumber
python master_hd.py ai-trich "%~1" --an-danh --model qwen3.6:35b --host http://localhost:11434 --out "ket_qua_AI\so_sanh_AI_%~n1.xlsx"
echo.
echo Ket qua: ket_qua_AI\so_sanh_AI_%~n1.xlsx
pause
