@echo off
REM Chay app kiem tra ho so tren Windows (can Python 3.10+)
cd /d %~dp0
python -m pip install -r requirements.txt
python -m streamlit run app.py
pause
