#!/bin/zsh
set -e
cd "$(dirname "$0")"
PYTHON="$(command -v python3 || true)"
if [[ -z "$PYTHON" ]]; then
  echo "未找到 Python 3，请先安装 Python 3 后重试。"
  read -k 1 "?按任意键关闭..."
  exit 1
fi
if [[ ! -d .venv ]]; then
  "$PYTHON" -m venv .venv
fi
if ! .venv/bin/python -c 'import streamlit, pandas, openpyxl, yaml' >/dev/null 2>&1; then
  .venv/bin/python -m pip install -r requirements.txt
fi
exec .venv/bin/python -m streamlit run app.py
