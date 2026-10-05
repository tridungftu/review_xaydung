@echo off
REM ================================================================
REM  Mo cua so cau hinh may Local AI trong mang LAN
REM  - Go IP may AI, bam "Kiem tra ket noi", chon model, bam "Luu"; hoac
REM  - Bam "Tu tim may AI trong mang" khi doi mang / may AI doi IP.
REM  Luu vao cau_hinh_may_ai.json (khong day len Git); moi cong cu doc tu file nay.
REM ================================================================
cd /d %~dp0
where pythonw >nul 2>nul
if not errorlevel 1 (
  start "" pythonw cau_hinh_may_ai_ui.py
  exit /b
)
where python >nul 2>nul
if errorlevel 1 (
  echo Chua cai Python. Tai tai https://www.python.org/downloads/ - nho tich "Add python.exe to PATH".
  pause
  exit /b
)
python cau_hinh_may_ai_ui.py
