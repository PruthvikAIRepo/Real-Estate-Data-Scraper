# PythonAnywhere Deployment Guide

## 📋 Prerequisites

1. **PythonAnywhere Account** - Sign up at https://www.pythonanywhere.com
2. **Free or Paid Plan** - Free tier works but has limitations

## 🚀 Deployment Steps

### 1. Upload Files to PythonAnywhere

**Option A: Git (Recommended)**
```bash
# On PythonAnywhere Bash console
git clone https://github.com/yourusername/Real-Estate-Data-Scraper.git
cd Real-Estate-Data-Scraper
```

**Option B: Manual Upload**
1. Go to **Files** tab
2. Upload all project files
3. Keep the same folder structure

### 2. Create Virtual Environment

```bash
# In PythonAnywhere Bash console
cd ~/Real-Estate-Data-Scraper
python3.10 -m venv env
source env/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Configure Web App

1. Go to **Web** tab
2. Click **Add a new web app**
3. Choose **Flask**
4. Python version: **Python 3.10**
5. Set path to: `/home/yourusername/Real-Estate-Data-Scraper/app.py`

### 4. Update WSGI Configuration

1. Click on **WSGI configuration file** link
2. Replace content with:

```python
import sys
import os

# Add your project directory to the sys.path
project_home = '/home/yourusername/Real-Estate-Data-Scraper'
if project_home not in sys.path:
    sys.path = [project_home] + sys.path

# Set environment variable
os.chdir(project_home)

# Import Flask app
from app import app as application
```

### 5. Set Virtual Environment

1. In **Web** tab, find **Virtualenv** section
2. Enter path: `/home/yourusername/Real-Estate-Data-Scraper/env`

### 6. Configure Static Files

Add these mappings in **Static files** section:

| URL | Directory |
|-----|-----------|
| /dashboard_data.json | /home/yourusername/Real-Estate-Data-Scraper/dashboard_data.json |
| /hemnet_listings.xlsx | /home/yourusername/Real-Estate-Data-Scraper/hemnet_listings.xlsx |
| /view_excel.html | /home/yourusername/Real-Estate-Data-Scraper/view_excel.html |

### 7. Reload Web App

Click **Reload yourusername.pythonanywhere.com** button

## 🎯 Test Your Deployment

### Access Dashboard
```
https://yourusername.pythonanywhere.com
```

### Test API Endpoints
```
https://yourusername.pythonanywhere.com/api/health
https://yourusername.pythonanywhere.com/api/scraper-status
```

### Run Scraper
1. Open dashboard
2. Click **"Run Scraper"** button
3. Watch progress in browser
4. Wait ~5 minutes for completion

## ⚙️ Configuration

### Update beautifulsoup.py

Change file paths to use absolute paths:

```python
# In beautifulsoup.py
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "hemnet_listings.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "scraper.log")
```

### Environment Variables (Optional)

Create `.env` file:
```
MAX_PAGES=51
MIN_DELAY=2
MAX_DELAY=4
```

## 🔧 Troubleshooting

### Error: Module not found
```bash
pip install -r requirements.txt
```

### Error: Permission denied
```bash
chmod +x app.py
```

### Scraper not running
Check logs:
```bash
tail -f ~/Real-Estate-Data-Scraper/scraper.log
```

### Dashboard not loading
Check error log:
- Web tab → Error log link
- Look for Python errors

## 📊 Monitoring

### View Logs
```bash
cd ~/Real-Estate-Data-Scraper
tail -f scraper.log
```

### Check Scraper Status
```bash
curl https://yourusername.pythonanywhere.com/api/scraper-status
```

### Manual Run (if button fails)
```bash
cd ~/Real-Estate-Data-Scraper
source env/bin/activate
python beautifulsoup.py
```

## 🔄 Updating Code

### Via Git
```bash
cd ~/Real-Estate-Data-Scraper
git pull origin main
```

### Reload App
Go to **Web** tab → Click **Reload**

## ⏰ Scheduled Tasks (Cron)

Set up automatic scraping:

1. Go to **Tasks** tab
2. Add scheduled task:
   ```bash
   cd /home/yourusername/Real-Estate-Data-Scraper && /home/yourusername/Real-Estate-Data-Scraper/env/bin/python beautifulsoup.py
   ```
3. Set time: **Daily at 3:00 AM (UTC)**

## 🎛️ Free Tier Limitations

- **100 seconds CPU time per day**
- **512MB storage**
- **Always-on tasks**: Not available (use scheduled tasks instead)
- **Scraper runtime**: ~5 minutes (may exceed free tier CPU limits)

**Recommendation**: Upgrade to paid plan ($5/month) for:
- Longer CPU time
- Background workers
- More reliable scraping

## 🆘 Support

- PythonAnywhere Help: https://help.pythonanywhere.com
- Dashboard Issues: Check browser console (F12)
- Scraper Issues: Check `scraper.log`

## ✅ Post-Deployment Checklist

- [ ] Dashboard loads at yourusername.pythonanywhere.com
- [ ] "Run Scraper" button works
- [ ] Progress updates show in real-time
- [ ] Excel file downloads work
- [ ] View Data page loads
- [ ] Scheduled task configured (optional)

## 🎉 You're Live!

Your Hemnet scraper dashboard is now accessible at:
```
https://yourusername.pythonanywhere.com
```

Click "Run Scraper" to start collecting data! 🚀
