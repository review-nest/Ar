from flask import Flask, render_template, request, send_file
from google_play_scraper import reviews, Sort, app as play_app
import threading
import requests
import json
import re
import time
import os
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import pandas as pd

# =====================================
# MOVIEPY SAFE IMPORT FIX (RENDER COMPATIBLE)
# =====================================
ImageClip = None
concatenate_videoclips = None

try:
    from moviepy.editor import ImageClip as IC, concatenate_videoclips as CVC
    ImageClip = IC
    concatenate_videoclips = CVC
except Exception:
    try:
        from moviepy.video.io.ImageClip import ImageClip as IC
        from moviepy.video.compositing.concatenate import concatenate_videoclips as CVC
        ImageClip = IC
        concatenate_videoclips = CVC
    except Exception as e:
        print("Warning: MoviePy module not loaded:", e)

app = Flask(__name__)

# =====================================
# CONFIG
# =====================================
SHEET_URL = "https://script.google.com/macros/s/AKfycbxYSKR7SAuMM7WuBXYH8PWQUYMD1vHSJ7OPlZhjRaJEbcDm3X-c8Q1M4UjBu9M3tJM4/exec"

MAX_FETCH = 50000
BATCH_SIZE = 300

# Global storage for fetched reviews and app info
CURRENT_FETCHED_REVIEWS = []
CURRENT_APP_INFO = {}
CURRENT_PACKAGE = ""
CURRENT_SEARCH_DATE = ""

# =====================================
# UTILITY: EXTRACT PACKAGE NAME FROM LINK/ID
# =====================================
def extract_package_id(input_str):
    if not input_str:
        return ""
    input_str = input_str.strip()
    match = re.search(r'id=([a-zA-Z0-9_.]+)', input_str)
    if match:
        return match.group(1)
    return input_str

# =====================================
# LOCAL HIERARCHICAL FOLDER STORAGE (UPDATED)
# Structure: Play Store / YYYY / MM / YYYY-MM-DD / YYYY-MM-DD_package.xlsx
# =====================================
def save_reviews_to_local_folder(package, search_date, reviews_data):
    if not reviews_data:
        return None
    try:
        # Parse search_date (e.g., '2026-08-02')
        dt = datetime.strptime(search_date, "%Y-%m-%d")
        
        year_folder = dt.strftime("%Y")         # e.g., '2026'
        month_folder = dt.strftime("%m")        # e.g., '08'
        date_folder = search_date               # e.g., '2026-08-02'
        
        # Structure: Play Store / Year / Month / Date
        base_dir = os.path.join("Play Store", year_folder, month_folder, date_folder)
        os.makedirs(base_dir, exist_ok=True)
        
        # File Name: 2026-08-02_com.example.app.xlsx
        file_name = f"{search_date}_{package}.xlsx"
        file_path = os.path.join(base_dir, file_name)
        
        rows = []
        for r in reviews_data:
            at = r.get("at")
            if hasattr(at, "strftime"):
                date_str = at.strftime("%Y-%m-%d")
                time_str = at.strftime("%H:%M:%S")
            else:
                date_str = str(at)[:10]
                time_str = str(at)[11:19]

            rows.append({
                "App Name": CURRENT_APP_INFO.get("title", package),
                "Username": str(r.get("userName", "")),
                "Review": str(r.get("content", "")),
                "Rating": int(r.get("score", 0)),
                "Date": date_str,
                "Time": time_str,
                "Package": package
            })
            
        df = pd.DataFrame(rows)
        df.to_excel(file_path, index=False)
        print(f"Locally saved to folder structure: {file_path}")
        return file_path
    except Exception as e:
        print("Local folder save error:", e)
        return None

# =====================================
# GOOGLE SHEET UPLOAD (BACKGROUND TASK)
# =====================================
def save_batch(package, search_date, rows):
    try:  
        response = requests.post(  
            SHEET_URL,  
            json={  
                "package": package,  
                "search_date": search_date,  
                "reviews": rows  
            },  
            timeout=120  
        )  
        print("Sheet Response:", response.text)  
    except Exception as e:  
        print("Batch Upload Error :", e)

def process_and_upload_async(package, search_date, reviews_data, app_title):
    rows = []  
    for r in reviews_data:  
        at = r.get("at")  
        if hasattr(at, "strftime"):  
            date = at.strftime("%Y-%m-%d")  
            review_time = at.strftime("%H:%M:%S")  
        else:  
            date = str(at)[:10]  
            review_time = str(at)[11:19]  

        rows.append({  
            "username": str(r.get("userName", "")),  
            "review": str(r.get("content", "")),  
            "rating": int(r.get("score", 0)),  
            "date": date,  
            "time": review_time,  
            "package": package  
        })  

    if not rows:  
        return  

    print(f"Async Task Started for {len(rows)} reviews...")  

    for i in range(0, len(rows), BATCH_SIZE):  
        batch = rows[i:i + BATCH_SIZE]  
        print(f"Uploading Batch {i + len(batch)}/{len(rows)}")  
        save_batch(package, search_date, batch)  
        time.sleep(0.3)  

    print("Google Sheet Upload Completed via Background Thread")
    save_reviews_to_local_folder(package, search_date, reviews_data)

# =====================================
# MATCH SYSTEM & KEYWORDS
# =====================================
def is_symbol_only(text):
    if not text:  
        return False  
    return all(not ch.isalnum() for ch in text)

def match_keyword(comment, keyword):
    comment = str(comment).strip()  
    keyword = str(keyword).strip()  

    if not comment or not keyword:  
        return False  

    if is_symbol_only(keyword):  
        m = re.search(r'([^\w\s]+)$', comment)  
        if not m:  
            return False  
        return m.group(1) == keyword  

    pattern = r'(?<!\w)' + re.escape(keyword.lower()) + r'(?!\w)'  
    return re.search(pattern, comment.lower()) is not None

def keyword_match(comment, keyword_text):
    if not keyword_text:  
        return True  

    keywords = [k.strip() for k in keyword_text.splitlines() if k.strip()]  

    if not keywords:  
        return True  

    for word in keywords:  
        if match_keyword(comment, word):  
            return True  

    return False

def review_pass(review, rating=None, keyword=None):
    if rating:  
        try:  
            if review.get("score") != int(rating):  
                return False  
        except:  
            return False  

    if keyword:  
        if not keyword_match(review.get("content", ""), keyword):  
            return False  

    return True

# =====================================
# APP INFO
# =====================================
def get_app_info(package):
    try:  
        info = play_app(package, country="in", lang="en")  
        return {  
            "title": info.get("title", package),  
            "icon": info.get("icon", ""),  
            "developer": info.get("developer", ""),  
            "installs": info.get("installs", ""),  
            "score": info.get("score", "")  
        }  
    except Exception as e:  
        print("App Info Error :", e)  
        return {  
            "title": package,  
            "icon": "",  
            "developer": "",  
            "installs": "",  
            "score": ""  
        }

def review_date(review):
    at = review.get("at")  
    if hasattr(at, "strftime"):  
        return at.strftime("%Y-%m-%d")  
    return str(at)[:10]  

# =====================================
# REVIEW FETCH ENGINE
# =====================================
def fetch_reviews(package, search_date, rating=None, keyword=None):
    data = []  
    token = None  
    total_scanned = 0  

    while True:  
        try:  
            result, token = reviews(  
                package,  
                lang="en",  
                country="in",  
                sort=Sort.NEWEST,  
                count=200,  
                continuation_token=token  
            )  
        except Exception as e:  
            print("Fetch Error :", e)  
            break  

        if not result:  
            break  

        stop = False  

        for review in result:  
            total_scanned += 1  
            r_date = review_date(review)  

            if r_date < search_date:  
                stop = True  
                break  

            if r_date != search_date:  
                continue  

            if not review_pass(review, rating, keyword):  
                continue  

            data.append(review)  

        print(f"Scanned : {total_scanned} | Matched : {len(data)}")  

        if stop or token is None or total_scanned >= MAX_FETCH:  
            break  

        time.sleep(0.1)  

    return data  

# =====================================
# EXCEL EXPORT ROUTE
# =====================================
@app.route("/export-excel")
def export_excel():
    global CURRENT_FETCHED_REVIEWS, CURRENT_PACKAGE, CURRENT_SEARCH_DATE
    if not CURRENT_FETCHED_REVIEWS:
        return "No reviews available to export.", 400
    
    file_path = save_reviews_to_local_folder(CURRENT_PACKAGE, CURRENT_SEARCH_DATE, CURRENT_FETCHED_REVIEWS)
    if file_path and os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "Error generating Excel file.", 500

# =====================================
# REEL VIEW ROUTE (SMOOTH AUTO-SCROLL)
# =====================================
@app.route("/reel")
def reel_view():
    global CURRENT_FETCHED_REVIEWS, CURRENT_APP_INFO
    return render_template(
        "reel.html",
        reviews=CURRENT_FETCHED_REVIEWS,
        app_info=CURRENT_APP_INFO
    )

# =====================================
# MAIN ROUTE
# =====================================
@app.route("/", methods=["GET", "POST"])
def home():
    global CURRENT_FETCHED_REVIEWS, CURRENT_APP_INFO, CURRENT_PACKAGE, CURRENT_SEARCH_DATE
    data = []  
    raw_input = ""  
    package = ""
    app_info = {}  

    if request.method == "POST":  
        raw_input = request.form.get("package", "").strip()  
        date = request.form.get("date", "").strip()  
        rating = request.form.get("rating", "").strip()  
        keyword = request.form.get("keyword", "").strip()  

        package = extract_package_id(raw_input)
        CURRENT_PACKAGE = package
        CURRENT_SEARCH_DATE = date

        if package:  
            app_info = get_app_info(package)  

        if package and date:  
            data = fetch_reviews(package=package, search_date=date, rating=rating, keyword=keyword)  

            CURRENT_FETCHED_REVIEWS = data
            CURRENT_APP_INFO = app_info

            if len(data) > 0:  
                thread = threading.Thread(
                    target=process_and_upload_async, 
                    args=(package, date, data, app_info.get("title", package))
                )
                thread.start()

    return render_template(  
        "index.html",  
        reviews=data,  
        package=raw_input,  
        app_info=app_info  
    )  

# =====================================
# HEALTH CHECK
# =====================================
@app.route("/health")
def health():
    return {  
        "status": "ok",  
        "service": "Google Play Review Fetcher"  
    }

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
