"""
Flask Backend for Hemnet Scraper Dashboard
Deployable to PythonAnywhere
"""

from flask import Flask, jsonify, send_from_directory, request
from flask_cors import CORS
import subprocess
import threading
import os
import sys
from datetime import datetime
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Global scraper status
scraper_status = {
    'running': False,
    'last_run': None,
    'last_run_time': None,
    'message': 'Ready to scrape',
    'progress': 0,
    'total_pages': 900,
    'current_municipality': '',
    'listings_added': 0
}

def run_scraper_background():
    """Run the scraper in background thread"""
    global scraper_status

    try:
        scraper_status['running'] = True
        scraper_status['message'] = 'Scraper started...'
        scraper_status['progress'] = 0

        logger.info("Starting scraper...")

        # Determine Python executable
        # On PythonAnywhere, sys.executable points to uwsgi, not python
        # So we need to find the correct python interpreter
        if 'uwsgi' in sys.executable.lower():
            # PythonAnywhere environment - use virtual env python
            base_dir = os.path.dirname(os.path.abspath(__file__))
            python_exe = os.path.join(base_dir, 'env', 'bin', 'python')
            if not os.path.exists(python_exe):
                # Fallback to system python3
                python_exe = 'python3'
        else:
            # Local development - use current python
            python_exe = sys.executable

        logger.info(f"Using Python: {python_exe}")

        # Set environment to force unbuffered output
        env = os.environ.copy()
        env['PYTHONUNBUFFERED'] = '1'

        # Run the scraper
        process = subprocess.Popen(
            [python_exe, '-u', 'beautifulsoup.py'],  # -u flag for unbuffered
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,  # Merge stderr into stdout to avoid blocking
            text=True,
            bufsize=0,  # Unbuffered
            env=env
        )

        # Monitor output line by line
        while True:
            line = process.stdout.readline()

            if not line:
                # No more output, check if process finished
                if process.poll() is not None:
                    break
                continue

            line = line.strip()
            if line:
                logger.info(f"Scraper output: {line}")

                # Update progress based on output
                if line.startswith("Municipality "):
                    try:
                        # e.g. "Municipality 12/290: Göteborgs kommun (1722 listings, 35 pages)"
                        parts = line.split(":")
                        m_progress = parts[0].replace("Municipality ", "").strip()
                        m_name = parts[1].strip().split("(")[0].strip()
                        current, total = m_progress.split("/")
                        scraper_status['current_municipality'] = m_name
                        scraper_status['message'] = f'Scraping {m_name} ({current}/{total} municipalities)...'
                    except Exception as e:
                        logger.debug(f"Could not parse municipality line: {e}")
                elif "Page" in line and ":" in line:
                    try:
                        # e.g. "Page 42: New: 30, Duplicates: 20"
                        parts = line.split()
                        for i, part in enumerate(parts):
                            if part == "Page" and i + 1 < len(parts):
                                page_num = int(parts[i + 1].rstrip(':'))
                                scraper_status['progress'] = page_num
                                break
                    except Exception as e:
                        logger.debug(f"Could not parse page number: {e}")

        # Wait for final completion
        return_code = process.wait()

        if return_code == 0:
            scraper_status['message'] = 'Scraper completed successfully!'
            scraper_status['last_run'] = 'Success'
            scraper_status['last_run_time'] = datetime.now().isoformat()
            logger.info("Scraper completed successfully")
        else:
            scraper_status['message'] = f'Scraper failed with exit code {return_code}'
            scraper_status['last_run'] = 'Failed'
            scraper_status['last_run_time'] = datetime.now().isoformat()
            logger.error(f"Scraper failed with exit code: {return_code}")

    except Exception as e:
        scraper_status['message'] = f'Error: {str(e)}'
        scraper_status['last_run'] = 'Error'
        scraper_status['last_run_time'] = datetime.now().isoformat()
        logger.error(f"Scraper error: {str(e)}")
    finally:
        scraper_status['running'] = False
        scraper_status['progress'] = scraper_status['total_pages']


# ========== ROUTES ==========

@app.route('/')
def index():
    """Serve dashboard"""
    return send_from_directory('.', 'dashboard.html')


@app.route('/<path:filename>')
def serve_static(filename):
    """Serve static files"""
    return send_from_directory('.', filename)


@app.route('/api/run-scraper', methods=['POST'])
def run_scraper():
    """API endpoint to start scraper"""
    global scraper_status

    if scraper_status['running']:
        return jsonify({
            'success': False,
            'message': 'Scraper is already running!',
            'status': scraper_status
        }), 400

    # Start scraper in background thread
    thread = threading.Thread(target=run_scraper_background, daemon=True)
    thread.start()

    logger.info("Scraper thread started")

    return jsonify({
        'success': True,
        'message': 'Scraper started successfully',
        'status': scraper_status
    })


@app.route('/api/scraper-status', methods=['GET'])
def get_status():
    """Get current scraper status"""
    return jsonify(scraper_status)


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat()
    })


# ========== ERROR HANDLERS ==========

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Not found'}), 404


@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Internal server error'}), 500


# ========== MAIN ==========

if __name__ == '__main__':
    print('='*60)
    print('Hemnet Scraper Dashboard Server')
    print('='*60)
    print('Dashboard: http://localhost:5000')
    print('API: http://localhost:5000/api')
    print('='*60)
    print('Press Ctrl+C to stop')
    print('='*60)

    # Run Flask app
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)
