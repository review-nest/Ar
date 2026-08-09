from flask import Flask, render_template, request, send_file
from google_play_scraper import reviews, Sort, app as play_app
import threading
import json
import re
import time
import os
import requests
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import pandas as pd


# =====================================
# MOVIEPY SAFE IMPORT FIX
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


# =====================================
# FLASK APP
# =====================================

app = Flask(__name__)


# =====================================
# CONFIG
# =====================================

SHEET_URL = "https://script.google.com/macros/s/AKfycbzH4wqpYp2JR9fUia9rCb-XOSxtCygLd98GhD_LKd3mQZwkoy6lttG5ToK8QRpGvZkDzQ/exec"

MAX_FETCH = 50000
BATCH_SIZE = 300


# =====================================
# GLOBAL STORAGE
# =====================================

CURRENT_FETCHED_REVIEWS = []
CURRENT_APP_INFO = {}
CURRENT_PACKAGE = ""
CURRENT_SEARCH_DATE = ""


# =====================================
# EXTRACT PACKAGE NAME
# =====================================

def extract_package_id(input_str):

    if not input_str:
        return ""

    input_str = input_str.strip()

    match = re.search(
        r'id=([a-zA-Z0-9_.]+)',
        input_str
    )

    if match:
        return match.group(1)

    return input_str


# =====================================
# SAFE FILE/FOLDER NAME
# =====================================

def safe_name(name):

    if not name:
        return "Unknown"

    name = str(name).strip()

    # Remove invalid Windows/Linux filename characters
    name = re.sub(
        r'[\\/:*?"<>|]+',
        '',
        name
    )

    # Remove extra spaces
    name = re.sub(
        r'\s+',
        ' ',
        name
    ).strip()

    if not name:
        return "Unknown"

    return name


# =====================================
# LOCAL HIERARCHICAL EXCEL STORAGE
#
# Play Store
#    └── Year
#         └── Month
#              └── Date
#                   └── AppName_Date.xlsx
# =====================================

def save_reviews_to_local_folder(
    package,
    search_date,
    reviews_data,
    app_title
):

    if not reviews_data:
        print("No reviews available for local Excel save.")
        return None

    try:

        # ---------------------------------
        # Convert search date
        # ---------------------------------

        dt = datetime.strptime(
            search_date,
            "%Y-%m-%d"
        )

        # ---------------------------------
        # Folder names
        # ---------------------------------

        year_folder = dt.strftime("%Y")
        month_folder = dt.strftime("%B")
        date_folder = dt.strftime("%Y-%m-%d")

        # ---------------------------------
        # Folder structure
        #
        # Play Store / Year / Month / Date
        # ---------------------------------

        base_dir = os.path.join(
            "Play Store",
            year_folder,
            month_folder,
            date_folder
        )

        os.makedirs(
            base_dir,
            exist_ok=True
        )

        # ---------------------------------
        # Safe App Name
        # ---------------------------------

        clean_app_name = safe_name(
            app_title
        )

        if not clean_app_name:
            clean_app_name = safe_name(
                package
            )

        # ---------------------------------
        # Excel filename
        # ---------------------------------

        file_name = (
            f"{clean_app_name}_{search_date}.xlsx"
        )

        file_path = os.path.join(
            base_dir,
            file_name
        )

        # ---------------------------------
        # Prepare Excel rows
        # ---------------------------------

        rows = []

        for r in reviews_data:

            at = r.get("at")

            # Review date/time
            if hasattr(at, "strftime"):

                review_date = at.strftime(
                    "%Y-%m-%d"
                )

                review_time = at.strftime(
                    "%H:%M:%S"
                )

            else:

                review_date = str(at)[:10]

                review_time = str(at)[11:19]

            rows.append({

                "Username": str(
                    r.get("userName", "")
                ),

                "Date": review_date,

                "Time": review_time,

                "Rating": int(
                    r.get("score", 0)
                ),

                "Review": str(
                    r.get("content", "")
                ),

                "Package": package,

                "App Name": app_title

            })

        # ---------------------------------
        # Create DataFrame
        # ---------------------------------

        df = pd.DataFrame(rows)

        # ---------------------------------
        # Save Excel
        # ---------------------------------

        df.to_excel(
            file_path,
            index=False
        )

        print(
            "Excel saved successfully:"
        )

        print(file_path)

        return file_path

    except Exception as e:

        print(
            "Local folder save error:",
            e
        )

        return None


# =====================================
# GOOGLE SHEET BATCH UPLOAD
# =====================================

def save_batch(
    package,
    search_date,
    rows
):

    try:

        payload = {

            "action": "save_reviews",

            "package": package,

            "search_date": search_date,

            "reviews": rows

        }

        response = requests.post(

            SHEET_URL,

            json=payload,

            timeout=120

        )

        print(
            "Google Sheet Status:",
            response.status_code
        )

        print(
            "Google Sheet Response:",
            response.text
        )

    except requests.exceptions.Timeout:

        print(
            "Google Sheet Upload Timeout"
        )

    except requests.exceptions.RequestException as e:

        print(
            "Google Sheet Request Error:",
            e
        )

    except Exception as e:

        print(
            "Google Sheet Upload Error:",
            e
        )


# =====================================
# BACKGROUND GOOGLE SHEET UPLOAD
# =====================================

def process_and_upload_async(
    package,
    search_date,
    reviews_data,
    app_title
):

    rows = []

    # ---------------------------------
    # Convert reviews
    # ---------------------------------

    for r in reviews_data:

        at = r.get("at")

        if hasattr(at, "strftime"):

            review_date_value = at.strftime(
                "%Y-%m-%d"
            )

            review_time = at.strftime(
                "%H:%M:%S"
            )

        else:

            review_date_value = str(at)[:10]

            review_time = str(at)[11:19]

        rows.append({

            "username": str(
                r.get("userName", "")
            ),

            "review": str(
                r.get("content", "")
            ),

            "rating": int(
                r.get("score", 0)
            ),

            "date": review_date_value,

            "time": review_time,

            "package": package,

            "app_name": app_title

        })

    # ---------------------------------
    # No data
    # ---------------------------------

    if not rows:

        print(
            "No reviews to upload."
        )

        return

    print(
        f"Async Upload Started: "
        f"{len(rows)} reviews"
    )

    # ---------------------------------
    # Google Sheet Batch Upload
    # ---------------------------------

    for i in range(
        0,
        len(rows),
        BATCH_SIZE
    ):

        batch = rows[
            i:i + BATCH_SIZE
        ]

        print(
            f"Uploading Batch: "
            f"{i + 1}-"
            f"{i + len(batch)} / "
            f"{len(rows)}"
        )

        save_batch(

            package,

            search_date,

            batch

        )

        time.sleep(0.3)

    print(
        "Google Sheet Upload Completed"
    )

    # ---------------------------------
    # Local Excel Backup
    # ---------------------------------

    save_reviews_to_local_folder(

        package,

        search_date,

        reviews_data,

        app_title

    )


# =====================================
# SYMBOL CHECK
# =====================================

def is_symbol_only(text):

    if not text:
        return False

    return all(
        not ch.isalnum()
        for ch in text
    )


# =====================================
# KEYWORD MATCH
# =====================================

def match_keyword(
    comment,
    keyword
):

    comment = str(
        comment
    ).strip()

    keyword = str(
        keyword
    ).strip()

    if not comment or not keyword:
        return False

    # Symbol-only keyword
    if is_symbol_only(keyword):

        m = re.search(
            r'([^\w\s]+)$',
            comment
        )

        if not m:
            return False

        return (
            m.group(1) == keyword
        )

    pattern = (
        r'(?<!\w)'
        + re.escape(
            keyword.lower()
        )
        + r'(?!\w)'
    )

    return (
        re.search(
            pattern,
            comment.lower()
        )
        is not None
    )


# =====================================
# MULTIPLE KEYWORDS
# =====================================

def keyword_match(
    comment,
    keyword_text
):

    if not keyword_text:
        return True

    keywords = [

        k.strip()

        for k in keyword_text.splitlines()

        if k.strip()

    ]

    if not keywords:
        return True

    for word in keywords:

        if match_keyword(
            comment,
            word
        ):

            return True

    return False


# =====================================
# REVIEW FILTER
# =====================================

def review_pass(
    review,
    rating=None,
    keyword=None
):

    # Rating filter
    if rating:

        try:

            if review.get(
                "score"
            ) != int(rating):

                return False

        except Exception:

            return False

    # Keyword filter
    if keyword:

        if not keyword_match(
            review.get(
                "content",
                ""
            ),
            keyword
        ):

            return False

    return True


# =====================================
# GET APP INFORMATION
# =====================================

def get_app_info(package):

    try:

        info = play_app(
            package,
            country="in",
            lang="en"
        )

        return {

            "title": info.get(
                "title",
                package
            ),

            "icon": info.get(
                "icon",
                ""
            ),

            "developer": info.get(
                "developer",
                ""
            ),

            "installs": info.get(
                "installs",
                ""
            ),

            "score": info.get(
                "score",
                ""
            )

        }

    except Exception as e:

        print(
            "App Info Error:",
            e
        )

        return {

            "title": package,

            "icon": "",

            "developer": "",

            "installs": "",

            "score": ""

        }


# =====================================
# REVIEW DATE
# =====================================

def review_date(review):

    at = review.get("at")

    if hasattr(at, "strftime"):

        return at.strftime(
            "%Y-%m-%d"
        )

    return str(at)[:10]


# =====================================
# REVIEW FETCH ENGINE
# =====================================

def fetch_reviews(
    package,
    search_date,
    rating=None,
    keyword=None
):

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

            print(
                "Fetch Error:",
                e
            )

            break

        if not result:
            break

        stop = False

        for review in result:

            total_scanned += 1

            r_date = review_date(
                review
            )

            # Stop when older than target date
            if r_date < search_date:

                stop = True

                break

            # Skip other dates
            if r_date != search_date:

                continue

            # Rating / keyword filter
            if not review_pass(
                review,
                rating,
                keyword
            ):

                continue

            data.append(review)

        print(
            f"Scanned: "
            f"{total_scanned} | "
            f"Matched: "
            f"{len(data)}"
        )

        if (
            stop
            or token is None
            or total_scanned >= MAX_FETCH
        ):

            break

        time.sleep(0.1)

    return data


# =====================================
# EXCEL EXPORT ROUTE
# =====================================

@app.route("/export-excel")
def export_excel():

    global CURRENT_FETCHED_REVIEWS
    global CURRENT_PACKAGE
    global CURRENT_SEARCH_DATE
    global CURRENT_APP_INFO

    if not CURRENT_FETCHED_REVIEWS:

        return (
            "No reviews available to export.",
            400
        )

    app_title = CURRENT_APP_INFO.get(
        "title",
        CURRENT_PACKAGE
    )

    file_path = save_reviews_to_local_folder(

        CURRENT_PACKAGE,

        CURRENT_SEARCH_DATE,

        CURRENT_FETCHED_REVIEWS,

        app_title

    )

    if (
        file_path
        and os.path.exists(file_path)
    ):

        return send_file(
            file_path,
            as_attachment=True
        )

    return (
        "Error generating Excel file.",
        500
    )


# =====================================
# REEL VIEW ROUTE
# =====================================

@app.route("/reel")
def reel_view():

    global CURRENT_FETCHED_REVIEWS
    global CURRENT_APP_INFO

    return render_template(

        "reel.html",

        reviews=CURRENT_FETCHED_REVIEWS,

        app_info=CURRENT_APP_INFO

    )


# =====================================
# MAIN ROUTE
# =====================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def home():

    global CURRENT_FETCHED_REVIEWS
    global CURRENT_APP_INFO
    global CURRENT_PACKAGE
    global CURRENT_SEARCH_DATE

    data = []

    raw_input = ""

    package = ""

    app_info = {}

    if request.method == "POST":

        # ---------------------------------
        # Form data
        # ---------------------------------

        raw_input = request.form.get(
            "package",
            ""
        ).strip()

        date = request.form.get(
            "date",
            ""
        ).strip()

        rating = request.form.get(
            "rating",
            ""
        ).strip()

        keyword = request.form.get(
            "keyword",
            ""
        ).strip()

        # ---------------------------------
        # Package ID
        # ---------------------------------

        package = extract_package_id(
            raw_input
        )

        CURRENT_PACKAGE = package

        CURRENT_SEARCH_DATE = date

        # ---------------------------------
        # App info
        # ---------------------------------

        if package:

            app_info = get_app_info(
                package
            )

        # ---------------------------------
        # Fetch reviews
        # ---------------------------------

        if package and date:

            data = fetch_reviews(

                package=package,

                search_date=date,

                rating=rating,

                keyword=keyword

            )

            CURRENT_FETCHED_REVIEWS = data

            CURRENT_APP_INFO = app_info

            # ---------------------------------
            # Background upload
            # ---------------------------------

            if len(data) > 0:

                thread = threading.Thread(

                    target=process_and_upload_async,

                    args=(

                        package,

                        date,

                        data,

                        app_info.get(
                            "title",
                            package
                        )

                    )

                )

                # Don't block Flask
                thread.daemon = True

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

        "service":
        "Google Play Review Fetcher"

    }


# =====================================
# RUN SERVER
# =====================================

if __name__ == "__main__":

    app.run(

        host="0.0.0.0",

        port=5000,

        debug=True

    )