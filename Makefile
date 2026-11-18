# Linux/macOS shortcuts. On Windows use scripts\run.ps1 and scripts\test.ps1.
PY ?= .venv/bin/python

.PHONY: venv install dev run test lint typecheck fmt

venv:
	python3.12 -m venv .venv

install: venv
	$(PY) -m pip install -r requirements.txt

dev: venv
	$(PY) -m pip install -r requirements-dev.txt

run:
	$(PY) -m streamlit run streamlit_app.py

lint:
	.venv/bin/ruff check planner app tests streamlit_app.py
	.venv/bin/ruff format --check planner app tests streamlit_app.py

fmt:
	.venv/bin/ruff format planner app tests streamlit_app.py

typecheck:
	.venv/bin/mypy

test: lint typecheck
	$(PY) -m pytest -s
