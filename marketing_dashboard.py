import logging
import re
import subprocess
import sys
from datetime import date, datetime
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, urljoin, urlparse

import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None


_engine = None
_db_read = None
_db_execute = None
_db_execute_many = None
_clean_text = None
_clean_number = None


def configure_marketing_dashboard(engine, db_read, db_execute, db_execute_many, clean_text, clean_number):
    global _engine, _db_read, _db_execute, _db_execute_many, _clean_text, _clean_number
    _engine = engine
    _db_read = db_read
    _db_execute = db_execute
    _db_execute_many = db_execute_many
    _clean_text = clean_text
    _clean_number = clean_number


def clean_text_value(value):
    if _clean_text:
        return _clean_text(value)
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").strip()


def clean_number_value(value):
    if _clean_number:
        return _clean_number(value)
    if pd.isna(value):
        return 0
    try:
        return float(str(value).replace(",", "").replace("₹", "").strip())
    except Exception:
        return 0


def execute_returning_id(query, params):
    with _engine.begin() as connection:
        return connection.execute(text(query), params or {}).scalar()


def ensure_marketing_tables():
    queries = [
        """
        CREATE TABLE IF NOT EXISTS marketing_pincode_master (
            id SERIAL PRIMARY KEY,
            pincode TEXT UNIQUE,
            city TEXT,
            state TEXT,
            zone TEXT,
            is_active TEXT DEFAULT 'Yes'
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS sku_master (
            id SERIAL PRIMARY KEY,
            fsn TEXT UNIQUE,
            title TEXT,
            brand TEXT,
            series TEXT,
            color TEXT,
            product_type TEXT,
            product_url TEXT,
            is_active TEXT DEFAULT 'Yes'
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS marketing_scrape_batches (
            id SERIAL PRIMARY KEY,
            run_datetime TEXT,
            keyword TEXT,
            run_scope TEXT,
            product_count INTEGER,
            pincode_count INTEGER,
            status TEXT,
            remark TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS marketing_rank_runs (
            id SERIAL PRIMARY KEY,
            batch_id INTEGER,
            run_datetime TEXT,
            keyword TEXT,
            selected_fsn TEXT,
            selected_sku_title TEXT,
            selected_brand TEXT,
            selected_series TEXT,
            selected_color TEXT,
            selected_type TEXT,
            pincode TEXT,
            city TEXT,
            state TEXT,
            my_rank INTEGER,
            my_price REAL,
            my_live_price_text TEXT,
            my_delivery_tat TEXT,
            stock_status TEXT,
            visibility_status TEXT
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS marketing_rank_products (
            id SERIAL PRIMARY KEY,
            run_id INTEGER,
            rank INTEGER,
            product_title TEXT,
            brand TEXT,
            price REAL,
            rating TEXT,
            review_count TEXT,
            delivery_tat TEXT,
            product_url TEXT,
            sponsored_status TEXT,
            flipkart_fsn TEXT,
            position_tag TEXT,
            is_my_sku BOOLEAN DEFAULT FALSE
        )
        """,
        "CREATE INDEX IF NOT EXISTS idx_marketing_rank_runs_batch ON marketing_rank_runs (batch_id)",
        "CREATE INDEX IF NOT EXISTS idx_marketing_rank_runs_filters ON marketing_rank_runs (keyword, selected_fsn, pincode)",
        "CREATE INDEX IF NOT EXISTS idx_marketing_rank_products_run ON marketing_rank_products (run_id)",
    ]

    for query in queries:
        _db_execute(query, clear_cache=False)


def get_existing_sku_master():
    ensure_marketing_tables()
    sku_df = _db_read(
        """
        SELECT fsn, title, brand, series, color, product_type, product_url
        FROM sku_master
        WHERE LOWER(COALESCE(is_active, 'Yes')) IN ('yes', 'y', '1', 'true', 'active')
        ORDER BY product_type, series, color, fsn
        """,
        use_cache=False,
    )

    if sku_df.empty:
        sku_df = _db_read(
            """
            SELECT
                fsn,
                MAX(title) AS title,
                '' AS brand,
                '' AS series,
                '' AS color,
                '' AS product_type,
                '' AS product_url
            FROM allocation_tracker
            WHERE fsn IS NOT NULL AND TRIM(fsn) <> ''
            GROUP BY fsn
            ORDER BY MAX(title), fsn
            """,
            use_cache=False,
        )

    for col in ["fsn", "title", "brand", "series", "color", "product_type", "product_url"]:
        if col not in sku_df.columns:
            sku_df[col] = ""
        sku_df[col] = sku_df[col].apply(clean_text_value)

    sku_df["display_name"] = sku_df.apply(make_sku_display_name, axis=1)
    return sku_df.drop_duplicates(subset=["fsn"])


def make_sku_display_name(row):
    parts = [
        clean_text_value(row.get("product_type", "")),
        clean_text_value(row.get("series", "")),
        clean_text_value(row.get("color", "")),
        clean_text_value(row.get("fsn", "")),
    ]
    display = " | ".join([part for part in parts if part])
    return display or clean_text_value(row.get("title", "")) or clean_text_value(row.get("fsn", ""))


def get_active_marketing_pincodes():
    ensure_marketing_tables()
    df = _db_read(
        """
        SELECT id, pincode, city, state, zone, is_active
        FROM marketing_pincode_master
        WHERE LOWER(COALESCE(is_active, 'Yes')) IN ('yes', 'y', '1', 'true', 'active')
        ORDER BY zone, state, city, pincode
        """,
        use_cache=False,
    )

    if not df.empty:
        df["pincode"] = df["pincode"].apply(clean_text_value)
    return df


def import_pincode_master(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = [clean_text_value(c).lower().replace(" ", "_") for c in df.columns]
    aliases = {
        "pin_code": "pincode",
        "pin": "pincode",
        "city_name": "city",
        "state_name": "state",
        "active": "is_active",
    }
    df = df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

    required = ["pincode", "city", "state"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing pincode master columns: {missing}")

    for col in ["pincode", "city", "state", "zone", "is_active"]:
        if col not in df.columns:
            df[col] = "Yes" if col == "is_active" else ""
        df[col] = df[col].apply(clean_text_value)

    rows = []
    for _, row in df.iterrows():
        pincode = clean_text_value(row.get("pincode"))
        if not pincode:
            continue
        rows.append(
            {
                "pincode": pincode,
                "city": clean_text_value(row.get("city")),
                "state": clean_text_value(row.get("state")),
                "zone": clean_text_value(row.get("zone")),
                "is_active": clean_text_value(row.get("is_active")) or "Yes",
            }
        )

    if not rows:
        return 0

    _db_execute_many(
        """
        INSERT INTO marketing_pincode_master (pincode, city, state, zone, is_active)
        VALUES (:pincode, :city, :state, :zone, :is_active)
        ON CONFLICT (pincode)
        DO UPDATE SET
            city = EXCLUDED.city,
            state = EXCLUDED.state,
            zone = EXCLUDED.zone,
            is_active = EXCLUDED.is_active
        """,
        rows,
    )
    return len(rows)


def create_marketing_scrape_batch(keyword, run_scope, product_count, pincode_count):
    return execute_returning_id(
        """
        INSERT INTO marketing_scrape_batches (
            run_datetime, keyword, run_scope, product_count, pincode_count, status, remark
        )
        VALUES (:run_datetime, :keyword, :run_scope, :product_count, :pincode_count, :status, :remark)
        RETURNING id
        """,
        {
            "run_datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "keyword": clean_text_value(keyword),
            "run_scope": clean_text_value(run_scope),
            "product_count": int(product_count),
            "pincode_count": int(pincode_count),
            "status": "Running",
            "remark": "",
        },
    )


def update_marketing_scrape_batch(batch_id, status, remark=""):
    if not batch_id:
        return
    _db_execute(
        """
        UPDATE marketing_scrape_batches
        SET status = :status, remark = :remark
        WHERE id = :id
        """,
        {"id": int(batch_id), "status": clean_text_value(status), "remark": clean_text_value(remark)},
    )


def extract_fsn_from_url(url):
    parsed = urlparse(clean_text_value(url))
    query = parse_qs(parsed.query)
    pid = query.get("pid", [""])[0]
    if pid:
        return clean_text_value(pid)
    match = re.search(r"(FAN[A-Z0-9]{10,})", clean_text_value(url), re.I)
    return match.group(1).upper() if match else ""


def price_to_number(price_text):
    text_value = clean_text_value(price_text)
    match = re.search(r"[\d,]+", text_value)
    return float(match.group(0).replace(",", "")) if match else None


def get_chrome_executable_path():
    candidates = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    return None


def launch_chromium(playwright, launch_args):
    try:
        return playwright.chromium.launch(**launch_args)
    except Exception as exc:
        if "Executable doesn't exist" not in str(exc) and "playwright install" not in str(exc).lower():
            raise

        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
            timeout=180,
        )
        return playwright.chromium.launch(**launch_args)


def set_flipkart_pincode(page, pincode):
    pincode = clean_text_value(pincode)
    if not pincode:
        return

    try:
        page.context.add_cookies(
            [
                {
                    "name": "LOCATION",
                    "value": pincode,
                    "domain": ".flipkart.com",
                    "path": "/",
                    "httpOnly": False,
                    "secure": True,
                    "sameSite": "Lax",
                }
            ]
        )
    except Exception:
        pass

    try:
        page.evaluate(
            """pin => {
                localStorage.setItem('pincode', pin);
                localStorage.setItem('location', pin);
            }""",
            pincode,
        )
    except Exception:
        pass


def scrape_flipkart_keyword(keyword, pincode, top_n=15):
    if sync_playwright is None:
        raise RuntimeError("Playwright is not installed in this Streamlit environment.")

    products = []
    keyword_url = f"https://www.flipkart.com/search?q={quote_plus(clean_text_value(keyword))}"
    chrome_path = get_chrome_executable_path()

    launch_args = {
        "headless": True,
        "args": ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
    }
    if chrome_path:
        launch_args["executable_path"] = chrome_path

    with sync_playwright() as playwright:
        browser = launch_chromium(playwright, launch_args)
        context = browser.new_context(
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        try:
            page.goto("https://www.flipkart.com", wait_until="domcontentloaded", timeout=45000)
            set_flipkart_pincode(page, pincode)
            page.goto(keyword_url, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3500)

            if "captcha" in page.content().lower():
                raise RuntimeError("Flipkart returned CAPTCHA in browser session.")

            cards = page.locator("a[href*='/p/'], a[href*='pid=']").element_handles()
            seen_urls = set()
            rank = 1

            for card in cards:
                if rank > top_n:
                    break
                try:
                    href = card.get_attribute("href") or ""
                    product_url = urljoin("https://www.flipkart.com", href)
                    fsn = extract_fsn_from_url(product_url)
                    if not product_url or product_url in seen_urls or not fsn:
                        continue

                    box = card.bounding_box()
                    if not box or box.get("width", 0) < 80 or box.get("height", 0) < 80:
                        continue

                    text_blob = clean_text_value(card.inner_text(timeout=2000))
                    if len(text_blob) < 20:
                        continue

                    lines = [line.strip() for line in text_blob.splitlines() if line.strip()]
                    title = lines[0] if lines else ""
                    price_line = next((line for line in lines if "₹" in line), "")
                    rating_line = next((line for line in lines if re.match(r"^[0-5](\.\d)?$", line)), "")
                    review_line = next((line for line in lines if "ratings" in line.lower() or "reviews" in line.lower()), "")
                    delivery_line = next(
                        (line for line in lines if any(word in line.lower() for word in ["delivery", "tomorrow", "today"])),
                        "",
                    )

                    products.append(
                        {
                            "rank": rank,
                            "product_title": title,
                            "brand": title.split()[0] if title else "",
                            "price": price_to_number(price_line),
                            "rating": rating_line,
                            "review_count": review_line,
                            "delivery_tat": delivery_line,
                            "product_url": product_url,
                            "sponsored_status": "Sponsored" if "sponsored" in text_blob.lower() else "Organic",
                            "flipkart_fsn": fsn,
                        }
                    )
                    seen_urls.add(product_url)
                    rank += 1
                except Exception:
                    continue
        finally:
            browser.close()

    return products


def title_similarity(a, b):
    return SequenceMatcher(None, clean_text_value(a).lower(), clean_text_value(b).lower()).ratio()


def identify_my_sku(products, selected_sku):
    selected_fsn = clean_text_value(selected_sku.get("fsn", "")).upper()
    selected_url = clean_text_value(selected_sku.get("product_url", ""))
    selected_title = clean_text_value(selected_sku.get("title", ""))
    selected_brand = clean_text_value(selected_sku.get("brand", "")).lower()

    for product in products:
        if selected_fsn and clean_text_value(product.get("flipkart_fsn", "")).upper() == selected_fsn:
            return product

    for product in products:
        if selected_url and selected_url in clean_text_value(product.get("product_url", "")):
            return product

    best_product = None
    best_score = 0
    for product in products:
        score = title_similarity(selected_title, product.get("product_title", ""))
        product_brand = clean_text_value(product.get("brand", "")).lower()
        if selected_brand and product_brand and selected_brand == product_brand:
            score += 0.1
        if score > best_score:
            best_score = score
            best_product = product

    return best_product if best_score >= 0.72 else None


def get_visibility_status(rank):
    if rank is None or pd.isna(rank):
        return "Not Visible"
    rank = int(rank)
    if rank <= 5:
        return "Strong"
    if rank <= 15:
        return "Average"
    return "Weak"


def extract_delivery_tat_from_text(text_blob):
    text_blob = clean_text_value(text_blob)
    patterns = [
        r"Delivery by [A-Za-z]+,\s*\d{1,2}\s+[A-Za-z]+",
        r"Delivery by [A-Za-z]+\s+\d{1,2}\s+[A-Za-z]+",
        r"Delivery by [A-Za-z]+",
        r"Usually delivered in \d+\s+days?",
        r"Get it by [A-Za-z]+,\s*\d{1,2}\s+[A-Za-z]+",
        r"Tomorrow|Today",
    ]
    for pattern in patterns:
        match = re.search(pattern, text_blob, re.I)
        if match:
            return clean_text_value(match.group(0))
    return ""


def scrape_flipkart_fsn_live_details_batch(sku_records, pincode, top15_products):
    details = {}
    top15_by_fsn = {
        clean_text_value(product.get("flipkart_fsn", "")).upper(): clean_text_value(product.get("product_url", ""))
        for product in top15_products
    }

    if sync_playwright is None:
        for sku in sku_records:
            details[clean_text_value(sku.get("fsn", ""))] = {
                "price": None,
                "live_price_text": "",
                "delivery_tat": "",
                "stock_status": "Playwright missing",
            }
        return details

    chrome_path = get_chrome_executable_path()
    launch_args = {
        "headless": True,
        "args": ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
    }
    if chrome_path:
        launch_args["executable_path"] = chrome_path

    with sync_playwright() as playwright:
        browser = launch_chromium(playwright, launch_args)
        context = browser.new_context(viewport={"width": 1366, "height": 900})
        page = context.new_page()

        try:
            page.goto("https://www.flipkart.com", wait_until="domcontentloaded", timeout=45000)
            set_flipkart_pincode(page, pincode)

            for sku in sku_records:
                fsn = clean_text_value(sku.get("fsn", "")).upper()
                product_url = top15_by_fsn.get(fsn) or clean_text_value(sku.get("product_url", ""))
                if not product_url and fsn:
                    product_url = f"https://www.flipkart.com/product/p/itm?pid={fsn}"

                result = {"price": None, "live_price_text": "", "delivery_tat": "", "stock_status": "Unknown"}
                try:
                    if not product_url:
                        result["stock_status"] = "URL Missing"
                    else:
                        page.goto(product_url, wait_until="domcontentloaded", timeout=45000)
                        page.wait_for_timeout(1800)
                        text_blob = clean_text_value(page.locator("body").inner_text(timeout=5000))
                        price_match = re.search(r"₹[\d,]+", text_blob)
                        if price_match:
                            result["live_price_text"] = price_match.group(0)
                            result["price"] = price_to_number(price_match.group(0))
                            result["stock_status"] = "In Stock"
                        strong_oos = any(
                            phrase in text_blob.lower()
                            for phrase in [
                                "currently out of stock",
                                "this item is currently out of stock",
                                "seller currently not available",
                            ]
                        )
                        if strong_oos and not result["live_price_text"]:
                            result["stock_status"] = "OOS"
                        result["delivery_tat"] = extract_delivery_tat_from_text(text_blob)
                except Exception as exc:
                    result["stock_status"] = f"Failed: {exc}"

                details[fsn] = result
        finally:
            browser.close()

    return details


def save_marketing_rank_run(keyword, selected_sku, pincode_row, products, my_product, batch_id, live_details):
    run_id = execute_returning_id(
        """
        INSERT INTO marketing_rank_runs (
            batch_id, run_datetime, keyword, selected_fsn, selected_sku_title, selected_brand,
            selected_series, selected_color, selected_type, pincode, city, state,
            my_rank, my_price, my_live_price_text, my_delivery_tat, stock_status, visibility_status
        )
        VALUES (
            :batch_id, :run_datetime, :keyword, :selected_fsn, :selected_sku_title, :selected_brand,
            :selected_series, :selected_color, :selected_type, :pincode, :city, :state,
            :my_rank, :my_price, :my_live_price_text, :my_delivery_tat, :stock_status, :visibility_status
        )
        RETURNING id
        """,
        {
            "batch_id": batch_id,
            "run_datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "keyword": clean_text_value(keyword),
            "selected_fsn": clean_text_value(selected_sku.get("fsn")),
            "selected_sku_title": clean_text_value(selected_sku.get("title")),
            "selected_brand": clean_text_value(selected_sku.get("brand")),
            "selected_series": clean_text_value(selected_sku.get("series")),
            "selected_color": clean_text_value(selected_sku.get("color")),
            "selected_type": clean_text_value(selected_sku.get("product_type")),
            "pincode": clean_text_value(pincode_row.get("pincode")),
            "city": clean_text_value(pincode_row.get("city")),
            "state": clean_text_value(pincode_row.get("state")),
            "my_rank": int(my_product.get("rank")) if my_product else None,
            "my_price": live_details.get("price"),
            "my_live_price_text": clean_text_value(live_details.get("live_price_text")),
            "my_delivery_tat": clean_text_value(live_details.get("delivery_tat")),
            "stock_status": clean_text_value(live_details.get("stock_status")),
            "visibility_status": get_visibility_status(my_product.get("rank") if my_product else None),
        },
    )

    selected_fsn = clean_text_value(selected_sku.get("fsn")).upper()
    product_rows = []
    my_rank = int(my_product.get("rank")) if my_product else None
    for product in products:
        rank = int(product.get("rank", 0))
        product_fsn = clean_text_value(product.get("flipkart_fsn", "")).upper()
        is_my_sku = bool(selected_fsn and product_fsn == selected_fsn)
        if is_my_sku:
            position_tag = "My SKU"
        elif my_rank and rank < my_rank:
            position_tag = "Above Me"
        elif my_rank and rank > my_rank:
            position_tag = "Below Me"
        else:
            position_tag = "Competitor"

        product_rows.append(
            {
                "run_id": run_id,
                "rank": rank,
                "product_title": clean_text_value(product.get("product_title")),
                "brand": clean_text_value(product.get("brand")),
                "price": product.get("price"),
                "rating": clean_text_value(product.get("rating")),
                "review_count": clean_text_value(product.get("review_count")),
                "delivery_tat": clean_text_value(product.get("delivery_tat")),
                "product_url": clean_text_value(product.get("product_url")),
                "sponsored_status": clean_text_value(product.get("sponsored_status")),
                "flipkart_fsn": product_fsn,
                "position_tag": position_tag,
                "is_my_sku": is_my_sku,
            }
        )

    if product_rows:
        _db_execute_many(
            """
            INSERT INTO marketing_rank_products (
                run_id, rank, product_title, brand, price, rating, review_count, delivery_tat,
                product_url, sponsored_status, flipkart_fsn, position_tag, is_my_sku
            )
            VALUES (
                :run_id, :rank, :product_title, :brand, :price, :rating, :review_count, :delivery_tat,
                :product_url, :sponsored_status, :flipkart_fsn, :position_tag, :is_my_sku
            )
            """,
            product_rows,
            clear_cache=False,
        )

    return run_id


def process_keyword_rank_check_all_products(keyword, sku_master, pincode_df, progress_bar=None, status_box=None):
    summary_rows = []
    product_rows_by_run = {}
    failures = []
    batch_id = None

    if pincode_df.empty:
        return pd.DataFrame(), product_rows_by_run, batch_id, ["No pincodes selected."]
    if sku_master.empty:
        return pd.DataFrame(), product_rows_by_run, batch_id, ["No SKU/FSN records selected."]

    sku_records = sku_master.to_dict("records")
    batch_id = create_marketing_scrape_batch(keyword, "Selected Products", len(sku_records), len(pincode_df))

    for pos, (_, pincode_row) in enumerate(pincode_df.iterrows(), start=1):
        pincode = clean_text_value(pincode_row.get("pincode"))
        label = f"{clean_text_value(pincode_row.get('city'))} - {pincode}"
        if status_box:
            status_box.info(f"Processing {label}: scraping Top 15, then checking live price/TAT for {len(sku_records)} FSNs")

        try:
            products = scrape_flipkart_keyword(keyword, pincode, top_n=15)
            live_details_by_fsn = scrape_flipkart_fsn_live_details_batch(sku_records, pincode, products)

            for sku_index, selected_sku in enumerate(sku_records):
                my_product = identify_my_sku(products, selected_sku)
                selected_fsn = clean_text_value(selected_sku.get("fsn"))
                live_details = live_details_by_fsn.get(
                    selected_fsn.upper(),
                    {"price": None, "live_price_text": "", "delivery_tat": "", "stock_status": "Unknown"},
                )
                products_to_save = products if sku_index == 0 else []
                run_id = save_marketing_rank_run(
                    keyword, selected_sku, pincode_row, products_to_save, my_product, batch_id, live_details
                )

                if sku_index == 0 and products:
                    product_rows_by_run[run_id] = pd.DataFrame(products)

                summary_rows.append(
                    {
                        "Run ID": run_id,
                        "Top 15 Run ID": run_id if sku_index == 0 else None,
                        "Keyword": keyword,
                        "Selected SKU/FSN": selected_fsn,
                        "Selected Product": make_sku_display_name(selected_sku),
                        "Type": selected_sku.get("product_type", ""),
                        "Series": selected_sku.get("series", ""),
                        "Color": selected_sku.get("color", ""),
                        "City": pincode_row.get("city", ""),
                        "State": pincode_row.get("state", ""),
                        "Pincode": pincode,
                        "My Rank": my_product.get("rank") if my_product else None,
                        "My Price": live_details.get("price"),
                        "My Live Price": live_details.get("live_price_text", ""),
                        "My Delivery TAT": live_details.get("delivery_tat", ""),
                        "Stock Status": live_details.get("stock_status", ""),
                        "Visibility Status": get_visibility_status(my_product.get("rank") if my_product else None),
                        "Match Status": "FSN matched in Top 15" if my_product else "FSN not in Top 15",
                    }
                )
        except Exception as exc:
            logging.exception("Marketing scrape failed for keyword=%s pincode=%s", keyword, pincode)
            failures.append(f"{label}: {exc}")
            for selected_sku in sku_records:
                summary_rows.append(
                    {
                        "Run ID": None,
                        "Top 15 Run ID": None,
                        "Keyword": keyword,
                        "Selected SKU/FSN": selected_sku.get("fsn", ""),
                        "Selected Product": make_sku_display_name(selected_sku),
                        "Type": selected_sku.get("product_type", ""),
                        "Series": selected_sku.get("series", ""),
                        "Color": selected_sku.get("color", ""),
                        "City": pincode_row.get("city", ""),
                        "State": pincode_row.get("state", ""),
                        "Pincode": pincode,
                        "My Rank": None,
                        "My Price": None,
                        "My Live Price": "",
                        "My Delivery TAT": "",
                        "Stock Status": "Scrape failed",
                        "Visibility Status": "Not Visible",
                        "Match Status": "Scrape failed",
                    }
                )

        if progress_bar:
            progress_bar.progress(pos / len(pincode_df))

    update_marketing_scrape_batch(batch_id, "Completed with failures" if failures else "Completed", "; ".join(failures[:5]))
    return pd.DataFrame(summary_rows), product_rows_by_run, batch_id, failures


def load_marketing_rank_products(run_id):
    if not run_id or pd.isna(run_id):
        return pd.DataFrame()
    return _db_read(
        """
        SELECT
            rank AS "Rank",
            product_title AS "Product Title",
            brand AS "Brand",
            price AS "Price",
            rating AS "Rating",
            review_count AS "Reviews",
            delivery_tat AS "Delivery TAT",
            sponsored_status AS "Sponsored/Organic",
            flipkart_fsn AS "Flipkart FSN",
            product_url AS "Product URL",
            position_tag AS "Position Tag"
        FROM marketing_rank_products
        WHERE run_id = :run_id
        ORDER BY rank
        """,
        {"run_id": int(run_id)},
        use_cache=False,
    )


def load_historical_rank_data(keyword=None, selected_fsn=None, pincode=None, start_date=None, end_date=None):
    where = []
    params = {}
    if keyword:
        where.append("keyword = :keyword")
        params["keyword"] = keyword
    if selected_fsn:
        where.append("selected_fsn = :selected_fsn")
        params["selected_fsn"] = selected_fsn
    if pincode:
        where.append("pincode = :pincode")
        params["pincode"] = pincode
    if start_date:
        where.append("DATE(run_datetime::timestamp) >= :start_date")
        params["start_date"] = str(start_date)
    if end_date:
        where.append("DATE(run_datetime::timestamp) <= :end_date")
        params["end_date"] = str(end_date)

    where_sql = "WHERE " + " AND ".join(where) if where else ""
    return _db_read(
        f"""
        SELECT
            id AS "Run ID",
            run_datetime,
            keyword AS "Keyword",
            selected_fsn AS "FSN",
            selected_sku_title AS "Title",
            selected_type AS "Type",
            selected_series AS "Series",
            selected_color AS "Color",
            pincode AS "Pincode",
            city AS "City",
            state AS "State",
            my_rank AS "Rank",
            my_price AS "Price",
            my_live_price_text AS "Live Price",
            my_delivery_tat AS "Delivery TAT",
            stock_status AS "Stock Status",
            visibility_status AS "Visibility"
        FROM marketing_rank_runs
        {where_sql}
        ORDER BY run_datetime DESC, id DESC
        LIMIT 5000
        """,
        params,
        use_cache=False,
    )


def get_marketing_scrape_batches():
    return _db_read(
        """
        SELECT id, run_datetime, keyword, run_scope, product_count, pincode_count, status, remark
        FROM marketing_scrape_batches
        ORDER BY id DESC
        LIMIT 200
        """,
        use_cache=False,
    )


def delete_marketing_batch(batch_id):
    _db_execute("DELETE FROM marketing_rank_products WHERE run_id IN (SELECT id FROM marketing_rank_runs WHERE batch_id = :id)", {"id": int(batch_id)}, clear_cache=False)
    _db_execute("DELETE FROM marketing_rank_runs WHERE batch_id = :id", {"id": int(batch_id)}, clear_cache=False)
    _db_execute("DELETE FROM marketing_scrape_batches WHERE id = :id", {"id": int(batch_id)})


def dataframes_to_excel(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def delivery_tat_to_days(value):
    value = clean_text_value(value).lower()
    if not value:
        return None
    if "today" in value:
        return 0
    if "tomorrow" in value:
        return 1
    match = re.search(r"(\d+)\s+days?", value)
    return int(match.group(1)) if match else None


def average_delivery_days(series):
    values = [delivery_tat_to_days(x) for x in series]
    values = [x for x in values if x is not None]
    return f"{sum(values) / len(values):.1f} days" if values else "NA"


def render_summary_filters(summary_df, prefix):
    filtered = summary_df.copy()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        types = st.multiselect(
            "Type",
            sorted([x for x in filtered["Type"].dropna().unique().tolist() if clean_text_value(x)]),
            key=f"{prefix}_type",
        )
    if types:
        filtered = filtered[filtered["Type"].isin(types)]

    with c2:
        series = st.multiselect(
            "Series",
            sorted([x for x in filtered["Series"].dropna().unique().tolist() if clean_text_value(x)]),
            key=f"{prefix}_series",
        )
    if series:
        filtered = filtered[filtered["Series"].isin(series)]

    with c3:
        colors = st.multiselect(
            "Color",
            sorted([x for x in filtered["Color"].dropna().unique().tolist() if clean_text_value(x)]),
            key=f"{prefix}_color",
        )
    if colors:
        filtered = filtered[filtered["Color"].isin(colors)]

    with c4:
        pincodes = st.multiselect(
            "Pincode",
            sorted([x for x in filtered["Pincode"].dropna().unique().tolist() if clean_text_value(x)]),
            key=f"{prefix}_pincode",
        )
    if pincodes:
        filtered = filtered[filtered["Pincode"].isin(pincodes)]

    return filtered


def render_competition_view(summary_df):
    if summary_df.empty:
        st.info("Run a keyword check to see competition visibility.")
        return

    filtered = render_summary_filters(summary_df, "competition")
    if filtered.empty:
        st.info("No rows match selected competition filters.")
        return

    filtered = filtered.copy()
    filtered["selector"] = filtered.apply(
        lambda row: f"{row['Selected Product']} | {row['Pincode']} | Rank {row['My Rank'] if pd.notna(row['My Rank']) else 'Not Visible'}",
        axis=1,
    )
    selected = st.selectbox("Select Product + Pincode", filtered["selector"].tolist())
    row = filtered[filtered["selector"] == selected].iloc[0]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Rank", int(row["My Rank"]) if pd.notna(row["My Rank"]) else "Not Visible")
    m2.metric("Live Price", row.get("My Live Price") or "NA")
    m3.metric("Delivery TAT", row.get("My Delivery TAT") or "NA")
    m4.metric("Stock", row.get("Stock Status") or "NA")

    top15_run_id = row.get("Top 15 Run ID") or row.get("Run ID")
    top15 = load_marketing_rank_products(top15_run_id)
    if top15.empty:
        st.info("Top 15 competition products are not available for this row.")
        return

    if pd.isna(row["My Rank"]):
        st.warning("Selected SKU not found in top 15 results for this keyword and pincode.")

    st.dataframe(top15, use_container_width=True)
    st.download_button(
        "Download Competition Products",
        data=dataframes_to_excel({"Competition": top15}),
        file_name="marketing_competition_products.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_history_tab(sku_master, active_pincodes):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        keyword = st.text_input("Keyword", key="history_keyword")
    with c2:
        sku_options = [""] + sku_master["fsn"].tolist()
        selected_fsn = st.selectbox("SKU/FSN", sku_options, key="history_fsn")
    with c3:
        pincode_options = [""] + active_pincodes["pincode"].tolist() if not active_pincodes.empty else [""]
        pincode = st.selectbox("Pincode", pincode_options, key="history_pincode")
    with c4:
        date_range = st.date_input("Date Range", value=(date.today(), date.today()), key="history_date_range")

    start_date, end_date = None, None
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range

    history = load_historical_rank_data(keyword, selected_fsn, pincode, start_date, end_date)
    if history.empty:
        st.info("No historical rank data found for selected filters.")
        return

    chart_df = history.copy()
    chart_df["run_datetime"] = pd.to_datetime(chart_df["run_datetime"], errors="coerce")
    chart_df["Rank"] = pd.to_numeric(chart_df["Rank"], errors="coerce")
    chart_df["Price"] = pd.to_numeric(chart_df["Price"], errors="coerce")
    chart_df["Delivery Days"] = chart_df["Delivery TAT"].apply(delivery_tat_to_days)
    visible = chart_df[chart_df["Rank"].notna()].sort_values("run_datetime")
    if not visible.empty:
        st.line_chart(visible, x="run_datetime", y="Rank")
        if visible["Price"].notna().any():
            st.line_chart(visible, x="run_datetime", y="Price")
        if visible["Delivery Days"].notna().any():
            st.line_chart(visible, x="run_datetime", y="Delivery Days")

    st.dataframe(history, use_container_width=True)
    st.download_button(
        "Download Historical Data",
        data=dataframes_to_excel({"History": history}),
        file_name="marketing_rank_history.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_manage_tab():
    batches = get_marketing_scrape_batches()
    if batches.empty:
        st.info("No previous marketing runs found.")
        return

    st.dataframe(batches, use_container_width=True)
    batch_id = st.selectbox("Select Run Batch to Delete", batches["id"].tolist())
    confirm = st.checkbox("I understand this will delete the selected marketing run.")
    if st.button("Delete Selected Run", disabled=not confirm):
        delete_marketing_batch(batch_id)
        st.success("Selected marketing run deleted.")
        st.rerun()


def show_marketing_dashboard(engine, db_read, db_execute, db_execute_many, clean_text, clean_number):
    configure_marketing_dashboard(engine, db_read, db_execute, db_execute_many, clean_text, clean_number)
    ensure_marketing_tables()

    sku_master = get_existing_sku_master()
    active_pincodes = get_active_marketing_pincodes()

    st.markdown("## Marketing Dashboard")
    st.caption("Track Flipkart keyword visibility, live price, delivery TAT, and Top 15 competition by pincode.")

    h1, h2, h3 = st.columns(3)
    h1.metric("Active Pincodes", len(active_pincodes))
    h2.metric("Available SKUs", len(sku_master))
    h3.metric("Marketplace", "Flipkart")

    tab_run, tab_summary, tab_competition, tab_history, tab_manage = st.tabs(
        ["Run Check", "Visibility Summary", "Competition View", "History", "Manage Runs"]
    )

    with tab_run:
        if active_pincodes.empty:
            st.warning("No active pincodes found. Upload Marketing Pincode Master before running rank checks.")
            uploaded_pincodes = st.file_uploader("Upload Marketing Pincode Master", type=["xlsx"], key="marketing_pincode_upload")
            if uploaded_pincodes and st.button("Import Pincode Master"):
                imported = import_pincode_master(uploaded_pincodes)
                st.success(f"Imported {imported} pincode rows.")
                st.rerun()

        if sku_master.empty:
            st.warning("No SKU/FSN records found. Add SKUs to sku_master or allocation_tracker first.")

        active_pincodes_for_run = active_pincodes.copy()
        if not active_pincodes_for_run.empty:
            active_pincodes_for_run["pincode_label"] = active_pincodes_for_run.apply(
                lambda row: f"{row.get('pincode', '')} - {row.get('city', '')}, {row.get('state', '')}",
                axis=1,
            )

        c1, c2, c3, c4 = st.columns([2.2, 1.0, 1.3, 2.6])
        with c1:
            keyword = st.text_input("Keyword", placeholder="Example: ceiling fan", key="marketing_keyword")
        with c2:
            run_for_all = st.checkbox("Run for all", value=True, key="marketing_run_for_all")
        with c3:
            pincode_scope = st.selectbox("Pincode Scope", ["Sample 5", "Selected pincodes", "All"], key="marketing_pincode_scope")

        pincode_labels = active_pincodes_for_run["pincode_label"].tolist() if not active_pincodes_for_run.empty else []
        with c4:
            selected_pincode_labels = st.multiselect(
                "Pincodes",
                pincode_labels,
                default=pincode_labels[:5],
                disabled=pincode_scope != "Selected pincodes" or active_pincodes_for_run.empty,
                key="marketing_selected_pincodes",
            )

        if pincode_scope == "Selected pincodes":
            selected_pincodes_df = active_pincodes_for_run[
                active_pincodes_for_run["pincode_label"].isin(selected_pincode_labels)
            ].copy()
        elif pincode_scope == "Sample 5":
            selected_pincodes_df = active_pincodes_for_run.head(5).copy()
        else:
            selected_pincodes_df = active_pincodes_for_run.copy()

        f1, f2, f3 = st.columns(3)
        type_options = sorted([x for x in sku_master["product_type"].dropna().unique().tolist() if clean_text_value(x)])
        with f1:
            selected_types = st.multiselect("Type", type_options, disabled=run_for_all or sku_master.empty)

        type_scope = sku_master.copy()
        if selected_types:
            type_scope = type_scope[type_scope["product_type"].isin(selected_types)]

        series_options = sorted([x for x in type_scope["series"].dropna().unique().tolist() if clean_text_value(x)])
        with f2:
            selected_series = st.multiselect("Series", series_options, disabled=run_for_all or sku_master.empty)

        series_scope = type_scope.copy()
        if selected_series:
            series_scope = series_scope[series_scope["series"].isin(selected_series)]

        color_options = sorted([x for x in series_scope["color"].dropna().unique().tolist() if clean_text_value(x)])
        with f3:
            selected_colors = st.multiselect("Color", color_options, disabled=run_for_all or sku_master.empty)

        sku_scope = sku_master.copy()
        if not run_for_all:
            if selected_types:
                sku_scope = sku_scope[sku_scope["product_type"].isin(selected_types)]
            if selected_series:
                sku_scope = sku_scope[sku_scope["series"].isin(selected_series)]
            if selected_colors:
                sku_scope = sku_scope[sku_scope["color"].isin(selected_colors)]

        a1, a2 = st.columns([4, 1])
        with a1:
            st.caption(
                f"Flow: scrape Flipkart Top 15 once per selected pincode ({len(selected_pincodes_df)}), "
                f"then match {len(sku_scope)} selected FSNs locally."
            )
        with a2:
            run_check = st.button("Run Rank Check", disabled=sku_scope.empty or selected_pincodes_df.empty)

        if run_check:
            if not clean_text_value(keyword):
                st.error("Please enter a keyword.")
            else:
                progress = st.progress(0)
                status = st.empty()
                with st.spinner("Scraping Flipkart ranking data..."):
                    summary_df, product_rows_by_run, batch_id, failures = process_keyword_rank_check_all_products(
                        keyword, sku_scope, selected_pincodes_df, progress_bar=progress, status_box=status
                    )
                st.session_state["marketing_summary_df"] = summary_df
                st.session_state["marketing_product_rows_by_run"] = product_rows_by_run
                st.session_state["marketing_batch_id"] = batch_id
                status.success("Rank check completed.")
                if failures:
                    with st.expander("Scraping failures"):
                        for failure in failures:
                            st.warning(failure)

    summary_df = st.session_state.get("marketing_summary_df", pd.DataFrame())

    with tab_summary:
        if summary_df.empty:
            st.info("Run a keyword check to see visibility summary.")
        else:
            filtered = render_summary_filters(summary_df, "summary")
            visible_ranks = pd.to_numeric(filtered["My Rank"], errors="coerce").dropna()
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Best Rank", int(visible_ranks.min()) if not visible_ranks.empty else "NA")
            m2.metric("Worst Rank", int(visible_ranks.max()) if not visible_ranks.empty else "NA")
            m3.metric("Average Rank", f"{visible_ranks.mean():.1f}" if not visible_ranks.empty else "NA")
            m4.metric("Rows Shown", len(filtered))
            m5.metric("Not Visible", int((filtered["Visibility Status"] == "Not Visible").sum()))
            m6.metric("Average Delivery TAT", average_delivery_days(filtered["My Delivery TAT"]))
            st.dataframe(filtered.drop(columns=["Run ID", "Top 15 Run ID"], errors="ignore"), use_container_width=True)
            st.download_button(
                "Download Filtered Summary",
                data=dataframes_to_excel({"Summary": filtered}),
                file_name="marketing_rank_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with tab_competition:
        render_competition_view(summary_df)

    with tab_history:
        render_history_tab(sku_master, active_pincodes)

    with tab_manage:
        render_manage_tab()
