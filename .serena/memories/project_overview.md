# Real-Estate-Data-Scraper - Project Overview

## Purpose
A production-ready Flask web scraper for **Hemnet.se** (Sweden's leading real estate platform) with a real-time dashboard, analytics, and mobile-responsive UI.

## Key Features
- Automated web scraping of up to 2,550+ listings (51 pages × 50 listings)
- Duplicate detection for deduplication
- Real-time progress tracking via Flask backend
- Background execution (non-blocking)
- Excel export (.xlsx) with styled headers
- Mobile-responsive dashboard with Chart.js analytics
- In-browser Excel viewer
- PythonAnywhere deployment support

## Tech Stack
- **Backend**: Python 3.10+, Flask 3.0.0, BeautifulSoup4, cloudscraper, openpyxl
- **Frontend**: HTML5, CSS3, JavaScript, Chart.js, SheetJS
- **Deployment**: PythonAnywhere

## Project Structure
```
Real-Estate-Data-Scraper/
├── app.py                      # Flask backend server (routes, background scraper)
├── beautifulsoup.py            # Main scraper with duplicate detection
├── generate_dashboard_data.py  # Analytics data generator
├── dashboard.html              # Main dashboard UI
├── view_excel.html             # Excel viewer page
├── requirements.txt            # Python dependencies
├── DEPLOY_PYTHONANYWHERE.md    # Deployment guide
├── README.md                   # Documentation
├── .gitignore                  # Git exclusions
├── hemnet_listings.xlsx        # Generated Excel data (runtime artifact)
└── dashboard_data.json         # Generated analytics (runtime artifact)
```

## API Endpoints
- `GET /` - Main dashboard
- `POST /api/run-scraper` - Start scraper in background
- `GET /api/scraper-status` - Current scraper status
- `GET /api/health` - Health check

## Key Configuration (beautifulsoup.py)
- `MAX_PAGES = 51`
- `MIN_DELAY = 2` (seconds)
- `MAX_DELAY = 4` (seconds)
- `REQUEST_TIMEOUT = 30`
- `EXCEL_FILE = 'hemnet_listings.xlsx'`
