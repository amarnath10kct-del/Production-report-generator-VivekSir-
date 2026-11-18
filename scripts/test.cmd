@echo off
rem Lint, type-check and run all tests on Windows: scripts\test.cmd
cd /d "%~dp0.."
if not exist .venv py -3.12 -m venv .venv || exit /b 1
.venv\Scripts\python.exe -m pip install -q -r requirements-dev.txt || exit /b 1
.venv\Scripts\ruff.exe check planner app tests streamlit_app.py || exit /b 1
.venv\Scripts\ruff.exe format --check planner app tests streamlit_app.py || exit /b 1
.venv\Scripts\mypy.exe || exit /b 1
.venv\Scripts\python.exe -m pytest -s %*
