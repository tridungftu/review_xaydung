@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d %~dp0
if not exist "ket_qua_CSDL" mkdir "ket_qua_CSDL"
REM Xuat co so du lieu ra Excel de theo doi (chi xem, sua tren Excel khong doi CSDL)
set /p HD=So hop dong (Enter = tat ca):
set /p NG=Tinh hieu luc den ngay dd/mm/yyyy (Enter = hom nay):
set ARGS=
if not "%NG%"=="" set ARGS=--ngay %NG%
python csdl_hop_dong.py xuat-excel %HD% %ARGS% --out "ket_qua_CSDL\CSDL_hop_dong.xlsx"
if errorlevel 1 (pause & exit /b)
start "" "ket_qua_CSDL\CSDL_hop_dong.xlsx"
pause
