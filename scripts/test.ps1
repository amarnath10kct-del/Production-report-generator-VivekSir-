# Lint, type-check and run all tests on Windows: .\scripts\test.ps1
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)
if (-not (Test-Path .venv)) {
    py -3.12 -m venv .venv
}
.\.venv\Scripts\python.exe -m pip install -q -r requirements-dev.txt
.\.venv\Scripts\ruff.exe check planner app tests streamlit_app.py
if (-not $?) { exit 1 }
.\.venv\Scripts\ruff.exe format --check planner app tests streamlit_app.py
if (-not $?) { exit 1 }
.\.venv\Scripts\mypy.exe
if (-not $?) { exit 1 }
.\.venv\Scripts\python.exe -m pytest -s @args
exit $LASTEXITCODE
