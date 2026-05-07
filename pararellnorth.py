# parallelnorth.py

import requests
import time
from datetime import datetime
from bs4 import BeautifulSoup

from dbconn import connect_to_db
from debug_utils import debug
from excel_reports import generate_changes_report

def get_new_session_id(conn, table_name):
    with conn.cursor() as cur:
        cur.execute(f"SELECT COALESCE(MAX(session_id), 0) + 1 FROM {table_name}")
        return cur.fetchone()[0]

def fetch_page_parallelnorth(url, retries=3, backoff_factor=0.5):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    for attempt in range(retries):
        debug(f"Fetching ParallelNorthIP page (attempt {attempt+1}): {url}", "DEBUG")
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                debug(f"Successfully fetched ParallelNorthIP page (status=200).", "DEBUG")
                return resp.text
            else:
                debug(f"ParallelNorth fetch failed: status={resp.status_code}, url={url}", "ERROR")
                return None
        except requests.exceptions.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            debug(f"Error fetching ParallelNorth page: {e}. Retrying in {wait_time}s...", "WARNING")
            time.sleep(wait_time)

    debug("❌ ParallelNorth fetch failed after max retries.", "ERROR")
    return None

def parse_parallelnorth(html_content, source_url):
    if not html_content:
        debug("No HTML content to parse for ParallelNorthIP.", "ERROR")
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    table_rows = soup.select(".table-portfolios .table-row")
    debug(f"Found {len(table_rows)} .table-row elements in the ParallelNorthIP HTML.", "DEBUG")

    results = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for row_idx, row in enumerate(table_rows, start=1):
        cells = row.select(".table-data")
        if len(cells) < 5:
            debug(f"Skipping row #{row_idx} — not enough columns (found {len(cells)}).", "DEBUG")
            continue

        sector = cells[0].get_text(strip=True)
        seller = cells[1].get_text(strip=True)
        sale_type = cells[2].get_text(strip=True)
        status = cells[3].get_text(strip=True)        # e.g. "Open" or "Closed"
        patent_count = cells[4].get_text(strip=True)  # e.g. "28" or "83 granted + 102 pending"

        record = {
            "sector": sector,
            "seller": seller,
            "sale_type": sale_type,
            "status": status,
            "patent_count": patent_count,
            "source_link": source_url,
            "added_date": run_date
        }
        results.append(record)

    debug(f"Parsed {len(results)} ParallelNorthIP portfolio entries.", "INFO")
    return results

def insert_parallelnorth_data(data, conn, session_id):
    table_name = "parallelnorth"
    debug(f"Inserting {len(data)} ParallelNorthIP records into DB (session_id={session_id})", "INFO")

    with conn.cursor() as cur:
        prev_session_id = session_id - 1

        debug(f"Fetching open items from prev_session_id={prev_session_id} to detect changes...", "DEBUG")
        cur.execute(f"""
            SELECT sector, seller, sale_type, status
            FROM {table_name}
            WHERE session_id = %s AND closed = FALSE
        """, (prev_session_id,))
        prev_open_rows = cur.fetchall()
        debug(f"Found {len(prev_open_rows)} open items in the previous session.", "DEBUG")

        prev_open_key_to_status = {}
        for (sect, sell, sale_t, stat) in prev_open_rows:
            key = (sect, sell, sale_t)
            prev_open_key_to_status[key] = stat

        current_keys = []
        for record in data:
            key = (record["sector"], record["seller"], record["sale_type"])
            current_keys.append(key)

            cur.execute(f"""
                SELECT COUNT(*) 
                FROM {table_name}
                WHERE sector = %s AND seller = %s AND sale_type = %s
            """, key)
            existed_before = (cur.fetchone()[0] > 0)

            newly_added = not existed_before
            closed = False

            old_status = prev_open_key_to_status.get(key)
            current_status = record["status"]
            status_changed = False

            if old_status and (old_status != current_status):
                status_changed = True

            cur.execute(f"""
                INSERT INTO {table_name}
                  (session_id, source_link, sector, seller, sale_type, status,
                   patent_count, added_date, newly_added, closed, status_changed)
                VALUES (%s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s)
            """, (
                session_id,
                record["source_link"],
                record["sector"],
                record["seller"],
                record["sale_type"],
                current_status,
                record["patent_count"],
                record["added_date"],
                newly_added,
                closed,
                status_changed
            ))
            debug(f"[INSERT] ParallelNorth {key} => new={newly_added}, closed={closed}, "
                  f"status_changed={status_changed}", "DEBUG")

        prev_open_keys = set(prev_open_key_to_status.keys())
        missing_keys = prev_open_keys - set(current_keys)
        debug(f"Found {len(missing_keys)} old items that disappeared (closed) in session {session_id}.", "DEBUG")

        for gone_key in missing_keys:
            (gone_sector, gone_seller, gone_sale_type) = gone_key
            old_stat = prev_open_key_to_status[gone_key]

            cur.execute(f"""
                INSERT INTO {table_name}
                  (session_id, source_link, sector, seller, sale_type, status,
                   patent_count, added_date, newly_added, closed, status_changed)
                SELECT 
                  %s AS session_id,
                  source_link,
                  sector,
                  seller,
                  sale_type,
                  status,
                  patent_count,
                  NOW(),
                  FALSE AS newly_added,
                  TRUE AS closed,
                  FALSE AS status_changed
                FROM {table_name}
                WHERE session_id = %s
                  AND sector = %s
                  AND seller = %s
                  AND sale_type = %s
                  AND closed = FALSE
                LIMIT 1
            """, (
                session_id,
                prev_session_id,
                gone_sector,
                gone_seller,
                gone_sale_type
            ))
            debug(f"[CLOSED] ParallelNorth item (sector={gone_sector}, seller={gone_seller}) "
                  f"in session {session_id}", "DEBUG")

        conn.commit()
        debug(f"✅ ParallelNorth commit complete for session_id={session_id}.", "INFO")

def run_parallelnorth_scraper():
    debug("Connecting to DB for ParallelNorthIP scraping...", "INFO")
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for ParallelNorthIP scraping. Aborting.", "ERROR")
        return

    try:
        table_name = "parallelnorth"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting ParallelNorthIP scrape for session_id={session_id}.", "INFO")

        url = "https://parallelnorthip.com/portfolios"
        html = fetch_page_parallelnorth(url)
        if not html:
            debug("No HTML returned from ParallelNorthIP site. Aborting scraper.", "ERROR")
            return

        data = parse_parallelnorth(html, url)
        if not data:
            debug("No data parsed from ParallelNorthIP. Insert aborted.", "WARNING")
            return

        insert_parallelnorth_data(data, conn, session_id)

        debug(f"Generating changes-only report for '{table_name}'...", "DEBUG")
        generate_changes_report(conn, table_name)

        debug(f"Finished ParallelNorthIP scraping for session_id={session_id}.", "INFO")
    finally:
        conn.close()
        debug("DB connection closed (ParallelNorth).", "DEBUG")


if __name__ == "__main__":
    run_parallelnorth_scraper()
