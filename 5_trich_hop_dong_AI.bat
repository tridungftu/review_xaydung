@echo off
chcp 65001 >nul
REM ================================================================
REM  TRICH THAM SO HOP DONG BANG LOCAL AI (may AI trong mang LAN)
REM  - Nhap dup: mo cua so.
REM  - Keo tha file / thu muc hop dong vao file nay: dua san vao danh sach.
REM  Cua so co 2 tab: "Trich hop dong" va "May AI" (nhap / tu tim IP may AI).
REM  Ket qua: ket_qua_AI\trich_lo\
REM ================================================================
cd /d %~dp0
where python >nul 2>nul
if errorlevel 1 goto chua_python
python -c "import openpyxl, docx, pdfplumber" >nul 2>nul
if errorlevel 1 goto cai_thu_vien
:da_co_thu_vien
python -c "import tkinterdnd2" >nul 2>nul
if errorlevel 1 python -m pip install -q tkinterdnd2 >nul 2>nul
where pythonw >nul 2>nul
if errorlevel 1 goto chay_python
start "" pythonw giao_dien_trich_hd.py %*
exit /b
:chay_python
python giao_dien_trich_hd.py %*
exit /b
:cai_thu_vien
echo Dang cai thu vien lan dau (can mang Internet)...
python -m pip install -q openpyxl python-docx pdfplumber pywin32
goto da_co_thu_vien
:chua_python
echo Chua cai Python. Tai tai https://www.python.org/downloads/
echo Khi cai nho tich o "Add python.exe to PATH", cai xong chay lai file nay.
pause
