@echo off
chcp 65001 >nul
REM ================================================================
REM  Day ma nguon + tai lieu chung len GitHub tridungftu/review_xaydung
REM  Chi cac file trong danh sach cua .gitignore duoc day len;
REM  hop dong, ho so, co so du lieu, ket qua KHONG bao gio bi day.
REM  Lan dau: Git se mo trinh duyet de dang nhap GitHub.
REM ================================================================
cd /d %~dp0
where git >nul 2>nul
if errorlevel 1 (
  echo Chua cai Git. Tai tai https://git-scm.com/download/win roi chay lai.
  pause
  exit /b
)
set LANDAU=0
if not exist ".git" (
  set LANDAU=1
  git init -b main
  git remote add origin https://github.com/tridungftu/review_xaydung.git
)
git config user.name >nul 2>nul || git config user.name "tridungftu"
git config user.email >nul 2>nul || git config user.email "tridungftu@googlemail.com"
git add -A
echo.
echo ===== Cac file se day len (kiem tra khong co du lieu hop dong) =====
git status --short
echo ====================================================================
set /p OK=Tiep tuc day len? (y/n):
if /i not "%OK%"=="y" exit /b
if "%LANDAU%"=="1" goto lan_dau
set /p MSG=Noi dung thay doi (Enter = "Cap nhat"):
if "%MSG%"=="" set MSG=Cap nhat
git commit -m "%MSG%"
goto day
:lan_dau
git commit -F commit_dau_tien.txt
if not errorlevel 1 del commit_dau_tien.txt
:day
git push -u origin main
pause
