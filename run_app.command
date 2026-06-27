#!/bin/zsh
set -e

cd "${0:A:h}"

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
if ! python -c "import streamlit, pandas, openpyxl, xlrd" >/dev/null 2>&1; then
  python -m pip install -r requirements.txt
fi
python -m streamlit run app.py
