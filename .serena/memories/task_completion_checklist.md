# Task Completion Checklist - Real-Estate-Data-Scraper

## After Code Changes
1. Ensure Flask server still starts: `python app.py`
2. Test relevant API endpoints manually if changed
3. If scraper changed, test with reduced `MAX_PAGES` (e.g., 1-2 pages)
4. Check `scraper.log` for any new error patterns
5. If dashboard HTML changed, verify in browser at `http://localhost:5000`

## Before Committing
1. Review changed files: `git diff`
2. Add specific files (avoid `git add .` to prevent committing large data files)
3. Key files to EXCLUDE from commits: `hemnet_listings.xlsx`, `dashboard_data.json`, `scraper.log`, `env/`
4. Write descriptive commit message

## No Automated Checks
- No test suite to run
- No linting to run
- No build step required

## Deployment
- After pushing to GitHub, reload web app on PythonAnywhere if deployed
- Check DEPLOY_PYTHONANYWHERE.md for deployment-specific steps
