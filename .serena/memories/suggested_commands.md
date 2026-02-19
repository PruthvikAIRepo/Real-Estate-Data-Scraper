# Suggested Commands - Real-Estate-Data-Scraper

## Environment Setup (Windows)
```bash
# Create virtual environment
python -m venv env

# Activate virtual environment (Windows)
env\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Running the Application
```bash
# Start Flask development server
python app.py
# Access at http://localhost:5000

# Run scraper directly (CLI)
python beautifulsoup.py

# Generate dashboard data from existing Excel
python generate_dashboard_data.py
```

## Git Operations
```bash
git status
git add <files>
git commit -m "message"
git push origin main
git pull origin main
git checkout -b feature/branch-name
```

## Debugging
- Check `scraper.log` for scraping errors
- Enable debug mode: set `app.run(debug=True)` in app.py
- Health check: `GET http://localhost:5000/api/health`
- Scraper status: `GET http://localhost:5000/api/scraper-status`

## Deployment (PythonAnywhere)
- See `DEPLOY_PYTHONANYWHERE.md` for full guide
- Python executable path: `/home/<username>/.virtualenvs/<venv>/bin/python3.10`
