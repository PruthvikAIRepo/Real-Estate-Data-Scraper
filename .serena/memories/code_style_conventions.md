# Code Style & Conventions - Real-Estate-Data-Scraper

## Python Style
- Python 3.10+ features used
- Module-level constants in UPPER_SNAKE_CASE (e.g., `MAX_PAGES`, `BASE_URL`)
- Functions in snake_case (e.g., `safe_get_text`, `load_existing_data`)
- Standard logging via `logging` module with `logger = logging.getLogger(__name__)`
- No type hints observed in existing code
- No docstrings in existing code
- Minimal inline comments

## Flask Patterns
- Standard Flask app setup with `Flask(__name__)`
- JSON API responses via `jsonify()`
- Background threading for long-running tasks
- Flask-CORS enabled for cross-origin requests

## Error Handling
- Try/except blocks around network requests and file operations
- Logging errors rather than raising exceptions in scraper
- Flask error handlers for 404 and 500

## File Naming
- Python files: snake_case (e.g., `beautifulsoup.py`, `generate_dashboard_data.py`)
- HTML files: snake_case (e.g., `dashboard.html`, `view_excel.html`)
- Data files: snake_case (e.g., `hemnet_listings.xlsx`, `dashboard_data.json`)

## No Testing Framework
- No test files or testing framework found in the project
- No linting configuration files (no `.flake8`, `.pylintrc`, `pyproject.toml`)
- No formatting configuration (no `black.toml`, `.isort.cfg`)
