# ipapproach.py

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

def fetch_page_ipapproach(url, retries=3, backoff_factor=0.5):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    for attempt in range(retries):
        debug(f"[IPApproach] Fetch attempt #{attempt+1}: {url}", "DEBUG")
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                debug(f"[IPApproach] success status=200 for {url}", "DEBUG")
                return resp.text
            else:
                debug(f"IPApproach fetch failed: status={resp.status_code}", "ERROR")
                return None
        except requests.exceptions.RequestException as ex:
            wait_time = backoff_factor * (2 ** attempt)
            debug(f"⚠️ IPApproach fetch error: {ex} => retry in {wait_time}s...", "WARNING")
            time.sleep(wait_time)
    debug("IPApproach fetch gave up after max retries.", "ERROR")
    return None

def parse_ipapproach_main(html_content, base_url):
    if not html_content:
        debug("No main HTML to parse for IPApproach main page.", "ERROR")
        return []
    soup = BeautifulSoup(html_content, "html.parser")

    cat_links = []
    p_tags = soup.find_all("p")
    for p in p_tags:
        if p.has_attr("style") and "text-align: center" in p["style"]:
            a_tags = p.find_all("a")
            for a in a_tags:
                href = a.get("href")
                cat_name = a.get_text(strip=True)
                if href and cat_name:
                    if href.startswith("http"):
                        full_link = href
                    else:
                        full_link = base_url.rstrip("/") + "/" + href.lstrip("/")
                    cat_links.append((cat_name, full_link))
    return cat_links

def parse_ipapproach_category(html_content, category_name):
    listings = []
    if not html_content:
        debug(f"No HTML content for category '{category_name}'.", "ERROR")
        return listings

    soup = BeautifulSoup(html_content, "html.parser")
    callout_divs = soup.find_all("div", class_="fl-callout-content")
    debug(f"[IPApproach] parse category='{category_name}', found {len(callout_divs)} callout divs.", "DEBUG")

    for cdiv in callout_divs:
        h2 = cdiv.find("h2", class_="fl-callout-title")
        if not h2:
            continue
        title = h2.get_text(strip=True)

        text_wrap = cdiv.find("div", class_="fl-callout-text")
        description = text_wrap.get_text("\n", strip=True) if text_wrap else ""

        patent_nums = re.findall(
            r'(?:[Uu]\.?[Ss]\.?\s*\d[\d,\.]*\s*[Bb]\d*|[EP]\s*\d[\d,\.]*|CA\s*\d[\d,\.]*)',
            description
        )
        patent_numbers_str = "\n".join(patent_nums)

        record = {
            "category": category_name,
            "title": title,
            "description": description,
            "patent_numbers": patent_numbers_str,
        }
        listings.append(record)

    return listings

def insert_ipapproach_data(data, conn, session_id):
    table_name = "ipapproach"
    debug(f"Inserting {len(data)} IPApproach listings => session_id={session_id}", "INFO")

    current_titles = [d["title"] for d in data]

    with conn.cursor() as cur:
        prev_session_id = session_id - 1
        cur.execute(f"""
            SELECT title FROM {table_name} 
            WHERE session_id=%s AND closed=FALSE
        """, (prev_session_id,))
        prev_open = [row[0] for row in cur.fetchall()]
        debug(f"[IPApproach] prev_session={prev_session_id}, found {len(prev_open)} open titles.", "DEBUG")

        for rec in data:
            title = rec["title"]
            cur.execute(f"SELECT COUNT(*) FROM {table_name} WHERE title=%s", (title,))
            existed = (cur.fetchone()[0] > 0)

            newly_added = not existed
            closed = False

            cur.execute(f"""
                INSERT INTO {table_name}
                  (session_id, category, title, description, patent_numbers,
                   newly_added, closed)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (
                session_id,
                rec["category"],
                rec["title"],
                rec["description"],
                rec["patent_numbers"],
                newly_added,
                closed
            ))
            debug(f"[INSERT] '{title}' => new={newly_added}", "DEBUG")

        closed_titles = [t for t in prev_open if t not in current_titles]
        debug(f"[IPApproach] found {len(closed_titles)} closed titles this run.", "DEBUG")

        for ct in closed_titles:
            cur.execute(f"""
                SELECT category, description, patent_numbers
                FROM {table_name}
                WHERE session_id = %s AND title = %s AND closed = FALSE
                LIMIT 1
            """, (prev_session_id, ct))
            row = cur.fetchone()
            if row:
                cat, desc, patnums = row
                cur.execute(f"""
                    INSERT INTO {table_name}
                      (session_id, category, title, description, patent_numbers,
                       newly_added, closed)
                    VALUES (%s, %s, %s, %s, %s,
                            %s, %s)
                """, (
                    session_id,
                    cat,
                    ct,
                    desc,
                    patnums,
                    False,
                    True,
                ))
                debug(f"[CLOSED] '{ct}' in session_id={session_id}", "DEBUG")

        conn.commit()
        debug("[IPApproach] insertion commit done.", "DEBUG")

def run_ipapproach_scraper():
    debug("[IPApproach] Connecting to database...", "INFO")
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for IPApproach scraping. Exiting.", "ERROR")
        return

    try:
        table_name = "ipapproach"
        session_id = get_new_session_id(conn, table_name)
        debug(f"[IPApproach] Starting scrape => session_id={session_id}", "INFO")

        main_url = "https://ipapproach.com/patents-for-sale/"
        main_html = fetch_page_ipapproach(main_url)
        if not main_html:
            debug("[IPApproach] No main HTML => aborting.", "ERROR")
            return

        cat_links = parse_ipapproach_main(main_html, "https://ipapproach.com")
        debug(f"[IPApproach] found {len(cat_links)} category links.", "INFO")

        all_listings = []
        for cat_name, cat_link in cat_links:
            debug(f"[IPApproach] scraping category='{cat_name}' => {cat_link}", "DEBUG")
            cat_html = fetch_page_ipapproach(cat_link)
            if cat_html:
                cat_listings = parse_ipapproach_category(cat_html, cat_name)
                debug(f"[IPApproach] category='{cat_name}' => {len(cat_listings)} listings", "DEBUG")
                all_listings.extend(cat_listings)
            else:
                debug(f"[IPApproach] skipping cat='{cat_name}' because no HTML", "WARNING")

        insert_ipapproach_data(all_listings, conn, session_id)
        
        debug(f"[IPApproach] generating changes-only Excel for '{table_name}'...", "DEBUG")
        generate_changes_report(conn, table_name)

        debug(f"[IPApproach] done. session_id={session_id}", "INFO")
    finally:
        debug("[IPApproach] closing DB connection...", "DEBUG")
        conn.close()
        debug("[IPApproach] DB connection closed.", "DEBUG")


if __name__ == "__main__":
    run_ipapproach_scraper()
