#!/bin/sh
# Lint, type-check and run all tests on Linux/macOS: ./scripts/test.sh
set -eu
cd "$(dirname "$0")/.."
[ -d .venv ] || python3.12 -m venv .venv
.venv/bin/pip install -q -r requirements-dev.txt
.venv/bin/ruff check planner app tests streamlit_app.py
.venv/bin/ruff format --check planner app tests streamlit_app.py
.venv/bin/mypy
exec .venv/bin/python -m pytest -s "$@"
