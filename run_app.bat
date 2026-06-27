@echo off
setlocal
cd /d "%~dp0"

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -c "import streamlit, pandas, openpyxl, xlrd" >nul 2>&1
if errorlevel 1 python -m pip install -r requirements.txt
python -m streamlit run app.py
