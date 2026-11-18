#!/bin/sh
# Run the app locally on Linux/macOS: ./scripts/run.sh  (http://localhost:8501)
set -eu
cd "$(dirname "$0")/.."
if [ ! -d .venv ]; then
  python3.12 -m venv .venv
  .venv/bin/pip install -r requirements.txt
fi
exec .venv/bin/python -m streamlit run streamlit_app.py --server.port "${PORT:-8501}"
