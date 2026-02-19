# Hemnet Real Estate Scraper

A production-ready Flask web scraper for **Hemnet.se** (Sweden's leading real estate platform) that scrapes **all ~41,000+ listings** across all 290 Swedish municipalities, with a real-time dashboard, analytics, and mobile-responsive UI.

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-3.0.0-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## How It Bypasses Hemnet's 2,500 Listing Cap

Hemnet limits all-Sweden searches to **50 pages × 50 listings = 2,500 listings** maximum, even though ~41,000+ exist. This scraper bypasses that limit by searching **municipality by municipality**.

```
All Sweden search  →  totalViewable: 2,500  (hard cap, useless)
Stockholm search   →  totalViewable: 2,140  (no cap, full access)
Göteborg search    →  totalViewable: 1,722  (no cap, full access)
...290 municipalities × no cap = all 41,000+ listings
```

No Swedish municipality exceeds 2,500 listings, so each municipality search returns all its listings with no cap applied.

---

## Features

### Core
- **Full Market Coverage** — Scrapes all ~41,000+ listings across all 290 Swedish municipalities
- **Live Page Count** — Queries Hemnet's GraphQL API at runtime to get the current listing count per municipality before scraping
- **Duplicate Detection** — O(1) set-based deduplication across the entire run (catches duplicates within and across municipalities)
- **Crash-Safe Checkpoints** — Saves Excel after every municipality so a crash never loses progress
- **Background Execution** — Scraper runs in a Flask background thread, non-blocking
- **Excel Export** — Formatted `.xlsx` with auto-styled headers

### Dashboard & Analytics
- **Real-time Progress** — Shows live page number + current municipality being scraped
- **Auto-Resume on Refresh** — Dashboard reconnects to a running scraper after page refresh
- **Live Dashboard Updates** — `dashboard_data.json` regenerated after every municipality so metrics update during scraping
- **Time Series Charts** — Track listings growth over time
- **Interactive Charts** — Chart.js powered analytics (pie, bar, line charts)
- **In-Browser Excel Viewer** — View data without downloading
- **Auto-Refresh** — Dashboard updates every 5 minutes

### Data Collected
- Property details (title, location, price, area, rooms, plot size)
- Listing metadata (type, label, broker name, broker logo)
- Images (up to 5 per listing)
- Tags and features (balcony, patio, etc.)
- Timestamps for time series analysis

---

## Architecture

```
municipalities.json          ← 290 Swedish municipality IDs (static, permanent)
        │
        ▼
beautifulsoup.py → scrape_hemnet()
        │
        ├─ Load municipalities.json (290 entries)
        ├─ Load existing Excel → build existing_ids set (O(1) dedup)
        │
        └─ For each of 290 municipalities:
               │
               ├─ get_live_count(id)        → POST GraphQL → returns live total
               ├─ pages = ceil(total / 50)  → calculated dynamically
               │
               ├─ For each page:
               │       fetch_page(?location_ids[]=ID&page=N)
               │       extract_listings(html, existing_ids)
               │       print("Page X: New: Y, Duplicates: Z")  ← Flask reads this
               │
               └─ Checkpoint:
                       save_to_excel()             ← crash-safe
                       generate_dashboard_data()   ← JSON updated live
```

---

## Tech Stack

**Backend:**
- Python 3.10+
- Flask 3.0.0
- BeautifulSoup4 (HTML parsing)
- cloudscraper (anti-bot bypass + GraphQL requests)
- openpyxl (Excel generation)

**Frontend:**
- HTML5, CSS3, JavaScript
- Chart.js (data visualization)
- SheetJS (client-side Excel parsing)

**Deployment:**
- PythonAnywhere compatible

---

## Installation

### Prerequisites
- Python 3.10 or higher
- pip
- Git

### Local Setup

1. **Clone the repository:**
```bash
git clone https://github.com/PruthvikAIRepo/Real-Estate-Data-Scraper.git
cd Real-Estate-Data-Scraper
```

2. **Create virtual environment:**
```bash
python -m venv env

# Windows
env\Scripts\activate

# macOS/Linux
source env/bin/activate
```

3. **Install dependencies:**
```bash
pip install -r requirements.txt
```

4. **Run the Flask server:**
```bash
python app.py
```

5. **Open in browser:**
```
http://localhost:5000
```

---

## Usage

### Via Dashboard (Recommended)
1. Open `http://localhost:5000`
2. Click **"Run Scraper"**
3. Watch live progress — button shows `Page 47 — Göteborgs kommun`
4. Dashboard metrics update automatically after each municipality
5. Full scrape takes **4–6 hours** for all 290 municipalities (~41,000 listings)

### Command Line
```bash
python beautifulsoup.py
```

### Quick Test (2–3 municipalities only)
Edit `municipalities.json` temporarily to test with a small subset:
```json
[
  {"id": "18031", "name": "Stockholms kommun"},
  {"id": "17920", "name": "Göteborgs kommun"}
]
```
Run, verify Excel and dashboard work, then restore the full file.

---

## Configuration

Edit `beautifulsoup.py` to customize:

```python
MIN_DELAY = 2           # Minimum delay between requests (seconds)
MAX_DELAY = 4           # Maximum delay between requests (seconds)
REQUEST_TIMEOUT = 30    # Request timeout (seconds)
EXCEL_FILE = "hemnet_listings.xlsx"
```

---

## Generated Files

| File | Description |
|------|-------------|
| `hemnet_listings.xlsx` | All scraped listings with formatted headers |
| `dashboard_data.json` | Analytics data regenerated after each municipality |
| `scraper.log` | Detailed per-page logs |

These files are excluded from git (see `.gitignore`) as they contain generated data.

---

## Project Structure

```
Real-Estate-Data-Scraper/
├── app.py                      # Flask backend — scraper API + status tracking
├── beautifulsoup.py            # Main scraper — municipality loop + deduplication
├── generate_dashboard_data.py  # Analytics generator — reads Excel, writes JSON
├── municipalities.json         # 290 Swedish municipality IDs (permanent, static)
├── dashboard.html              # Main dashboard UI
├── view_excel.html             # In-browser Excel viewer
├── requirements.txt            # Python dependencies
├── DEPLOY_PYTHONANYWHERE.md    # PythonAnywhere deployment guide
└── README.md
```

---

## API Endpoints

### `GET /`
Returns the main dashboard HTML.

### `POST /api/run-scraper`
Starts the scraper in a background thread.
```json
{
  "success": true,
  "message": "Scraper started successfully",
  "status": { ... }
}
```

### `GET /api/scraper-status`
Returns current scraper status including live progress.
```json
{
  "running": true,
  "progress": 47,
  "current_municipality": "Göteborgs kommun",
  "message": "Scraping Göteborgs kommun (7/290 municipalities)...",
  "last_run": null,
  "last_run_time": null
}
```

### `GET /api/health`
Health check.
```json
{
  "status": "healthy",
  "timestamp": "2026-02-20T12:00:00"
}
```

---

## Dashboard Features

### Live Progress (During Scraping)
- Button shows: `Page 47 — Göteborgs kommun`
- Auto-resumes live tracking after page refresh
- Dashboard metrics update after each municipality completes

### Analytics
- Total listings count, average/median price, price per m²
- Time series — listings growth over scraping sessions
- Property type distribution (pie chart)
- Price distribution (bar chart)
- Rooms distribution
- Top 10 locations by listing count
- Top 10 brokers by market share
- Features analysis (balcony, patio, etc.)

---

## Performance

| Metric | Value |
|--------|-------|
| Total listings scraped | ~41,000+ |
| Municipalities covered | 290 |
| Scraping time | ~4–6 hours |
| Request delay | 2–4s (polite rate limiting) |
| Deduplication | O(1) set lookup |
| Checkpoint frequency | After every municipality |
| Memory usage | ~300MB during scraping |

---

## Deployment

### PythonAnywhere
Follow the guide in **[DEPLOY_PYTHONANYWHERE.md](DEPLOY_PYTHONANYWHERE.md)**

**Quick steps:**
1. Create PythonAnywhere account (free tier available)
2. Clone repository via Bash console
3. Set up virtual environment and install dependencies
4. Configure Flask web app and WSGI file
5. Go live at `yourusername.pythonanywhere.com`

---

## Troubleshooting

**Scraper stuck at a municipality**
Check `scraper.log`. Usually a network timeout. Re-run — deduplication skips already-scraped listings automatically.

**Dashboard shows stale count**
Click **"Refresh Now"** button. JSON updates after each municipality, not every page.

**Button resets after page refresh**
Fixed in current version — dashboard auto-detects a running scraper and resumes live tracking.

**Excel file not found**
Run the scraper at least once to generate `hemnet_listings.xlsx`.

**No listings found on a page**
Hemnet may have changed their HTML structure. Check CSS selectors in `extract_listings()` in `beautifulsoup.py`.

---

## municipalities.json

Contains all 290 Swedish municipalities with their Hemnet internal IDs. This file is **permanent** — Swedish municipalities are fixed by law and their Hemnet IDs never change. It never needs to be updated.

```json
[
  {"id": "18031", "name": "Stockholms kommun"},
  {"id": "17920", "name": "Göteborgs kommun"},
  {"id": "17989", "name": "Malmö kommun"},
  ...
]
```

These are **municipality (Kommun)** level IDs — Sweden's middle administrative level (below County/Län, above District/Område). Chosen because:
- Fixed count: exactly 290
- No municipality exceeds 2,500 listings (bypasses Hemnet cap)
- Complete coverage: all listings in Sweden belong to exactly one municipality

---

## Disclaimer

This scraper is for **educational purposes only**. Please:
- Review Hemnet.se's Terms of Service before use
- Use built-in rate limiting (2–4 second delays between requests)
- Use responsibly and do not overload their servers

---

## Author

**Pruthvik**
- GitHub: [@PruthvikAIRepo](https://github.com/PruthvikAIRepo)
- Email: pruthvikchandarana9081@gmail.com

---

## Acknowledgments

- [Hemnet.se](https://www.hemnet.se) — Data source
- [Flask](https://flask.palletsprojects.com/) — Web framework
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) — HTML parsing
- [Chart.js](https://www.chartjs.org/) — Data visualization
- [cloudscraper](https://github.com/VeNoMouS/cloudscraper) — Anti-bot bypass
