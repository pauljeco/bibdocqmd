PYTHON ?= python3
VENV := .venv

.PHONY: setup setup-python setup-addin test test-python test-addin validate-manifest check

setup: setup-python setup-addin

setup-python:
	$(PYTHON) -m venv $(VENV)
	$(VENV)/bin/python -m pip install --upgrade pip
	$(VENV)/bin/python -m pip install -e '.[dev]'

setup-addin:
	cd office-addin && npm ci

test: test-python test-addin

test-python:
	$(VENV)/bin/python -m pytest -q

test-addin:
	cd office-addin && npm test

validate-manifest:
	cd office-addin && npm run validate

check: test-python
	cd office-addin && npm run check
