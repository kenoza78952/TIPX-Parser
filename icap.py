# icap.py

import requests
import time
import re
from datetime import datetime
from bs4 import BeautifulSoup

from dbconn import connect_to_db
from debug_utils import debug
from excel_reports import generate_changes_report

def get_new_session_id(conn, table_name):
    with conn.cursor() as cur:
        cur.execute(f"SELECT COALESCE(MAX(session_id), 0) + 1 FROM {table_name}")
        return cur.fetchone()[0]

def fetch_page_icap(url, retries=3, backoff_factor=0.5):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    for attempt in range(retries):
        debug(f"Attempt #{attempt+1} to fetch ICAP page: {url}", "DEBUG")
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                debug(f"✅ Success: ICAP page fetched (status=200).", "DEBUG")
                return resp.text
            else:
                debug(f"❌ ICAP fetch failed: status={resp.status_code}, url={url}", "ERROR")
                return None
        except requests.exceptions.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            debug(f"⚠️ Error fetching ICAP page: {e}. Retrying in {wait_time}s...", "WARNING")
            time.sleep(wait_time)

    debug("❌ ICAP fetch failed after max retries.", "ERROR")
    return None

def parse_icap(html_content, source_url):
    if not html_content:
        debug("No ICAP HTML content to parse.", "ERROR")
        return []

    soup = BeautifulSoup(html_content, "html.parser")

    debug("Parsing ICAP HTML to find all <li> that contain <strong>...", "DEBUG")
    all_li = soup.find_all("li")
    debug(f"Found {len(all_li)} <li> elements total.", "DEBUG")

    results = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for li_item in all_li:
        strong_tag = li_item.find("strong")
        if not strong_tag:
            continue

        title = strong_tag.get_text(strip=True)

        sub_ul = li_item.find("ul")
        bullet_points = []
        if sub_ul:
            sub_li_list = sub_ul.find_all("li", recursive=True)
            for sub_li in sub_li_list:
                bullet_text = sub_li.get_text(strip=True)
                bullet_points.append(bullet_text)

        description = "\n".join(bullet_points)

        record = {
            "title": title,
            "description": description,
            "added_date": run_date,
            "source_link": source_url,
        }
        debug(f"Parsed ICAP listing: '{title}', {len(bullet_points)} bullet items", "DEBUG")
        results.append(record)

    debug(f"Total ICAP listings with <strong>: {len(results)}", "INFO")
    return results

def insert_icap_data(data, conn, session_id):
    table_name = "icap"
    debug(f"Inserting {len(data)} ICAP listings into DB (session_id={session_id})", "INFO")

    with conn.cursor() as cur:
        prev_session_id = session_id - 1
        debug(f"Fetching open items from prev_session_id={prev_session_id}", "DEBUG")
        cur.execute(
            f"SELECT title FROM {table_name} WHERE session_id = %s AND closed = FALSE",
            (prev_session_id,)
        )
        prev_open_titles = [r[0] for r in cur.fetchall()]
        debug(f"Previous open titles count: {len(prev_open_titles)}", "DEBUG")

        current_titles = [d["title"] for d in data]

        for record in data:
            cur.execute(
                f"SELECT COUNT(*) FROM {table_name} WHERE title = %s",
                (record["title"],)
            )
            existed_before = (cur.fetchone()[0] > 0)

            newly_added = not existed_before
            closed = False

            cur.execute(f"""
                INSERT INTO {table_name}
                (session_id, source_link, title, description,
                 added_date, newly_added, closed)
                VALUES (%s, %s, %s, %s,
                        %s, %s, %s)
            """, (
                session_id,
                record["source_link"],
                record["title"],
                record["description"],
                record["added_date"],
                newly_added,
                closed
            ))
            debug(f"[INSERT] ICAP title='{record['title']}' => newly_added={newly_added}, closed={closed}", "DEBUG")

        closed_titles = [t for t in prev_open_titles if t not in current_titles]
        debug(f"Found {len(closed_titles)} newly 'closed' ICAP titles in this scrape.", "DEBUG")

        for ctitle in closed_titles:
            cur.execute(f"""
                SELECT COUNT(*)
                FROM {table_name}
                WHERE session_id = %s AND title = %s AND closed = TRUE
            """, (session_id, ctitle))
            already_closed = cur.fetchone()[0]
            if already_closed > 0:
                debug(f"Title='{ctitle}' is already marked closed in session {session_id}. Skipping...", "DEBUG")
                continue

            cur.execute(f"""
                SELECT description, source_link
                FROM {table_name}
                WHERE session_id = %s AND title = %s AND closed = FALSE
                LIMIT 1
            """, (prev_session_id, ctitle))
            row = cur.fetchone()
            if row:
                desc, src = row
                cur.execute(f"""
                    INSERT INTO {table_name}
                    (session_id, source_link, title, description,
                     added_date, newly_added, closed)
                    VALUES (%s, %s, %s, %s,
                            NOW(), %s, %s)
                """, (
                    session_id,
                    src,
                    ctitle,
                    desc,
                    False,
                    True
                ))
                debug(f"[CLOSED] ICAP title='{ctitle}' in session {session_id}", "DEBUG")

        conn.commit()
        debug(f"ICAP commit complete for session_id={session_id}.", "INFO")

def run_icap_scraper():
    debug("Connecting to database for ICAP scraping...", "INFO")
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for ICAP scraping. Exiting.", "ERROR")
        return

    try:
        table_name = "icap"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting ICAP scrape for session_id={session_id}.", "INFO")

        icap_url = "https://icappatentbrokerage.com/patents-available" 
        html = fetch_page_icap(icap_url)
        if not html:
            debug("No HTML content from ICAP. Aborting scraper.", "ERROR")
            return

        data = parse_icap(html, icap_url)
        if not data:
            debug("No ICAP data parsed from the page. No insertion performed.", "WARNING")
            return

        insert_icap_data(data, conn, session_id)

        debug(f"Generating changes-only report for '{table_name}'...", "DEBUG")
        generate_changes_report(conn, table_name)

        debug(f"Finished ICAP scraping for session_id={session_id}.", "INFO")
    finally:
        debug("Closing DB connection for ICAP scraping...", "DEBUG")
        conn.close()

if __name__ == "__main__":
    run_icap_scraper()
