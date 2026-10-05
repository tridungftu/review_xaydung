#!/usr/bin/env bash
# Chạy app kiểm tra hồ sơ trên macOS / Linux (cần Python 3.10+)
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
