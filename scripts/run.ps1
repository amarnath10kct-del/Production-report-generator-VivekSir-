# Run the app locally on Windows: .\scripts\run.ps1  (http://localhost:8501)
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)
if (-not (Test-Path .venv)) {
    py -3.12 -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
}
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py --server.port 8501
