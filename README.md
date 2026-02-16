# 🏠 Hemnet Real Estate Scraper

A production-ready Flask web scraper for **Hemnet.se** (Sweden's leading real estate platform) with a real-time dashboard, analytics, and mobile-responsive UI.

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-3.0.0-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📋 Features

### 🔥 Core Features
- **Automated Web Scraping** - Scrapes up to 2,550+ listings (51 pages × 50 listings)
- **Duplicate Detection** - Smart deduplication prevents duplicate entries
- **Real-time Progress Tracking** - Live updates during scraping with progress bar
- **Background Execution** - Non-blocking scraper runs via Flask backend
- **Excel Export** - Formatted `.xlsx` files with auto-styled headers

### 📊 Dashboard & Analytics
- **Mobile-Responsive Design** - Perfect on desktop, tablet, and mobile
- **Time Series Visualization** - Track listings growth over time
- **Interactive Charts** - Chart.js powered analytics (pie, bar, line charts)
- **In-Browser Excel Viewer** - View data without downloading Excel
- **Auto-Refresh** - Dashboard updates every 5 minutes
- **Real-time Notifications** - Success/error alerts with toast messages

### 🎯 Data Collected
- Property details (title, location, price, area, rooms)
- Listing metadata (type, label, broker info)
- Images (up to 5 per listing)
- Tags and features (balcony, patio, etc.)
- Timestamps for time series analysis

---

## 🖼️ Screenshots

### Dashboard Overview
![Dashboard](https://via.placeholder.com/800x400?text=Dashboard+Screenshot)

### Mobile View
![Mobile](https://via.placeholder.com/400x800?text=Mobile+View)

### Excel Viewer
![Excel Viewer](https://via.placeholder.com/800x400?text=Excel+Viewer)

---

## 🛠️ Tech Stack

**Backend:**
- Python 3.10+
- Flask 3.0.0 (Web framework)
- BeautifulSoup4 (HTML parsing)
- cloudscraper (Anti-bot bypass)
- openpyxl (Excel generation)

**Frontend:**
- HTML5, CSS3, JavaScript
- Chart.js (Data visualization)
- SheetJS (Client-side Excel parsing)

**Deployment:**
- PythonAnywhere compatible
- Git-based deployment

---

## 📦 Installation

### Prerequisites
- Python 3.10 or higher
- pip (Python package manager)
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

## 🚀 Usage

### Running the Scraper

**Option 1: Via Dashboard (Recommended)**
1. Open http://localhost:5000 in your browser
2. Click the **"Run Scraper"** button
3. Watch real-time progress (0/51 → 51/51)
4. Wait ~5 minutes for completion
5. View results in dashboard or download Excel

**Option 2: Command Line**
```bash
python beautifulsoup.py
```

### Configuration

Edit `beautifulsoup.py` to customize:

```python
MAX_PAGES = 51          # Number of pages to scrape (max 51)
MIN_DELAY = 2           # Minimum delay between requests (seconds)
MAX_DELAY = 4           # Maximum delay between requests (seconds)
REQUEST_TIMEOUT = 30    # Request timeout (seconds)
```

### Generated Files

- **`hemnet_listings.xlsx`** - Main data file with all listings
- **`dashboard_data.json`** - Analytics data for dashboard
- **`scraper.log`** - Detailed scraping logs

---

## 📊 Dashboard Features

### Metrics Display
- Total listings count
- New vs existing listings
- Listing type distribution (Bostad, Nybyggnadsprojekt)
- Premium/Plus/Max/Standard label breakdown

### Charts
1. **Time Series Total** - Listings growth over time
2. **Time Series Types** - Listing type trends
3. **Listing Type Distribution** - Pie chart
4. **Listing Label Distribution** - Pie chart
5. **Top 10 Locations** - Bar chart

### Actions
- **Refresh Data** - Reload dashboard analytics
- **Run Scraper** - Start new scraping session
- **View Excel** - Open in-browser Excel viewer
- **Download Excel** - Download `.xlsx` file

---

## 🌐 Deployment

### PythonAnywhere Deployment

Follow the comprehensive guide in **[DEPLOY_PYTHONANYWHERE.md](DEPLOY_PYTHONANYWHERE.md)**

**Quick steps:**
1. Create PythonAnywhere account (free tier available)
2. Clone repository via Bash console
3. Set up virtual environment
4. Configure Flask web app
5. Update WSGI configuration
6. Go live at `yourusername.pythonanywhere.com`

### Deployment Checklist
- [ ] Python 3.10 virtual environment created
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] WSGI file configured for Flask
- [ ] Static files mapped correctly
- [ ] Web app reloaded
- [ ] Dashboard accessible at public URL

---

## 📁 Project Structure

```
Real-Estate-Data-Scraper/
├── app.py                      # Flask backend server
├── beautifulsoup.py            # Main scraper with duplicate detection
├── generate_dashboard_data.py  # Analytics data generator
├── dashboard.html              # Main dashboard UI
├── view_excel.html             # Excel viewer page
├── requirements.txt            # Python dependencies
├── DEPLOY_PYTHONANYWHERE.md    # Deployment guide
├── README.md                   # This file
├── .gitignore                  # Git exclusions
├── hemnet_listings.xlsx        # Generated Excel data
└── dashboard_data.json         # Generated analytics
```

---

## 🔧 API Endpoints

### `GET /`
Returns the main dashboard HTML

### `POST /api/run-scraper`
Starts the scraper in background
```json
Response: {
  "success": true,
  "message": "Scraper started successfully",
  "status": { ... }
}
```

### `GET /api/scraper-status`
Returns current scraper status
```json
Response: {
  "running": false,
  "progress": 51,
  "total_pages": 51,
  "message": "Scraper completed successfully!",
  "last_run": "Success",
  "last_run_time": "2026-02-17T04:20:00"
}
```

### `GET /api/health`
Health check endpoint
```json
Response: {
  "status": "healthy",
  "timestamp": "2026-02-17T04:20:00"
}
```

---

## ⚙️ Configuration

### Environment Variables (Optional)

Create a `.env` file:
```env
MAX_PAGES=51
MIN_DELAY=2
MAX_DELAY=4
FLASK_DEBUG=False
```

### Logging

Logs are written to `scraper.log` with format:
```
2026-02-17 04:20:00 - INFO - Page 1: New: 50, Duplicates: 0
```

---

## 🐛 Troubleshooting

### Common Issues

**Issue:** Scraper stuck at page X
- **Solution:** Check `scraper.log` for errors. May be network timeout or rate limiting.

**Issue:** Dashboard not loading
- **Solution:** Ensure Flask server is running (`python app.py`)

**Issue:** No new listings found
- **Solution:** Normal behavior if all listings already scraped. Run again later for new data.

**Issue:** Excel file not found
- **Solution:** Run scraper at least once to generate `hemnet_listings.xlsx`

### Debug Mode

Enable Flask debug mode for development:
```python
app.run(debug=True)
```

---

## 📈 Performance

- **Scraping Speed:** ~5 minutes for 51 pages (2,550 listings)
- **Memory Usage:** ~200MB during scraping
- **Storage:** ~5MB per 2,500 listings (Excel + JSON)
- **Dashboard Load:** <2 seconds with 10,000+ listings

---

## 🤝 Contributing

Contributions are welcome! Here's how:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## ⚠️ Disclaimer

This scraper is for **educational purposes only**. Please:
- Review Hemnet.se's Terms of Service before use
- Implement rate limiting (built-in: 2-4 second delays)
- Use responsibly and ethically
- Do not overload their servers

---

## 👤 Author

**Pruthvik**
- GitHub: [@PruthvikAIRepo](https://github.com/PruthvikAIRepo)
- Email: pruthvikchandarana9081@gmail.com

---

## 🙏 Acknowledgments

- [Hemnet.se](https://www.hemnet.se) - Data source
- [Flask](https://flask.palletsprojects.com/) - Web framework
- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) - HTML parsing
- [Chart.js](https://www.chartjs.org/) - Data visualization
- [PythonAnywhere](https://www.pythonanywhere.com/) - Hosting platform

---

## 📞 Support

If you have questions or need help:
1. Check [DEPLOY_PYTHONANYWHERE.md](DEPLOY_PYTHONANYWHERE.md) for deployment issues
2. Review `scraper.log` for error details
3. Open an issue on GitHub
4. Contact via email

---

## 🔮 Future Enhancements

- [ ] PostgreSQL database support
- [ ] Historical data comparison
- [ ] Email notifications on new listings
- [ ] Price drop alerts
- [ ] Multi-city support
- [ ] RESTful API for third-party integrations
- [ ] Docker containerization
- [ ] Scheduled automatic scraping (cron jobs)

---

<div align="center">

**⭐ Star this repo if you find it useful!**

Made with ❤️ using Python & Flask

</div>
