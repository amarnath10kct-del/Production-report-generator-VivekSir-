@echo off
rem Run the app locally on Windows: scripts\run.cmd  (http://localhost:8501)
cd /d "%~dp0.."
if not exist .venv (
  py -3.12 -m venv .venv || exit /b 1
  .venv\Scripts\python.exe -m pip install -r requirements.txt || exit /b 1
)
.venv\Scripts\python.exe -m streamlit run streamlit_app.py --server.port 8501
