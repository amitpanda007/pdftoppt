.PHONY: help install install-dev test lint format clean build upload upload-test

help:  ## Show this help message
	@echo "Available commands:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

install:  ## Install package
	pip install -e .

install-dev:  ## Install package with development dependencies
	pip install -e ".[dev]"

test:  ## Run tests
	python -m pytest tests/ -v

test-coverage:  ## Run tests with coverage
	python -m pytest tests/ --cov=pdftoppt --cov-report=html --cov-report=term

lint:  ## Run linting
	flake8 pdftoppt/
	mypy pdftoppt/

format:  ## Format code
	black pdftoppt/ tests/ examples.py
	isort pdftoppt/ tests/ examples.py

clean:  ## Clean build artifacts
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf htmlcov/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

build:  ## Build package
	python -m build

upload-test:  ## Upload to TestPyPI
	python -m twine upload --repository testpypi dist/*

upload:  ## Upload to PyPI
	python -m twine upload dist/*

check:  ## Check package before upload
	python -m twine check dist/*

all-checks: format lint test  ## Run all checks (format, lint, test)

release-prep: clean build check  ## Prepare for release (clean, build, check)
