@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
if not exist "ket_qua_CSDL" mkdir "ket_qua_CSDL"
REM Tao phieu nhap phu luc cho mot hop dong (co san BOQ va tham so hien hanh de tra ma)
set /p HD=Nhap so hop dong (vd 0404):
python csdl_hop_dong.py tao-phieu "%HD%" --out "ket_qua_CSDL\phieu_phu_luc_%HD%.xlsx"
if errorlevel 1 (pause & exit /b)
start "" "ket_qua_CSDL\phieu_phu_luc_%HD%.xlsx"
pause
