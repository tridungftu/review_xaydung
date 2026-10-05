@echo off
chcp 65001 >nul
REM ================================================================
REM  Trich 36 tham so cho CA LO hop dong, nho may Local AI trong LAN
REM  Cach dung:
REM   - KEO THA mot thu muc (hoac nhieu file hop dong) vao file nay; hoac
REM   - Chep hop dong (.doc/.docx/.pdf) vao thu muc hop_dong_can_trich roi bam dup file nay.
REM  Ket qua: ket_qua_AI\trich_lo\  (master tung HD, so sanh AI, file TONG_HOP)
REM  Chay lai: hop dong da co ket qua AI se khong hoi lai (them --lam-lai de hoi lai).
REM ================================================================
cd /d %~dp0
where python >nul 2>nul
if errorlevel 1 (
  echo Chua cai Python. Tai tai https://www.python.org/downloads/ - nho tich "Add python.exe to PATH".
  pause
  exit /b
)
python -m pip install -q openpyxl python-docx pdfplumber pywin32 >nul 2>nul
if not exist "hop_dong_can_trich" mkdir "hop_dong_can_trich"
if "%~1"=="" goto mac_dinh
python trich_lo_hop_dong.py chay %*
goto xong
:mac_dinh
python trich_lo_hop_dong.py chay "hop_dong_can_trich"
:xong
if exist "ket_qua_AI\trich_lo" start "" "ket_qua_AI\trich_lo"
echo.
pause
