# rzv.py
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

def fetch_page_rzv(url, retries=3, backoff_factor=0.3):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    for attempt in range(retries):
        debug(f"Attempt #{attempt+1} to fetch RZV page: {url}", "DEBUG")
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                debug(f"Success: RZV page fetched (status=200).", "DEBUG")
                return resp.text
            else:
                debug(f"RZV fetch failed: status={resp.status_code}, url={url}", "ERROR")
                return None
        except requests.exceptions.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            debug(f"Error fetching RZV page: {e}. Retrying in {wait_time}s...", "WARNING")
            time.sleep(wait_time)

    debug("RZV fetch failed after max retries.", "ERROR")
    return None

def parse_rzv_opportunities(html_content, source_url):
    """
    Parses the RZV HTML to extract each portfolio listing.
    Logs details and returns a list of dicts.
    """
    if not html_content:
        debug("No RZV HTML content to parse.", "ERROR")
        return []

    debug("Parsing RZV HTML to locate <h2> + <p> blocks...", "DEBUG")
    soup = BeautifulSoup(html_content, "html.parser")
    main_div = soup.find("div", class_="8u skel-cell-important")
    if not main_div:
        debug("Could not locate main <div class='8u skel-cell-important'> for RZV data.", "ERROR")
        return []

    h2_list = main_div.find_all("h2")
    debug(f"Found {len(h2_list)} <h2> elements in RZV main_div.", "DEBUG")

    results = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for h2 in h2_list:
        title = h2.get_text(strip=True)
        debug(f"Processing <h2> title: '{title}'", "DEBUG")

        p_tag = h2.find_next_sibling("p")
        if not p_tag:
            debug(f"H2 title '{title}' but no <p> found. Skipping...", "WARNING")
            continue

        desc = p_tag.get_text(strip=True, separator=" ")

        possible_status = None
        if "(Licenses Signed)" in title or "(Licenses Signed)" in desc:
            possible_status = "Licenses Signed"
        elif "(Received Offer)" in title or "(Received Offer)" in desc:
            possible_status = "Received Offer"

        families_count = None
        patents_count = None

        def find_number_in_parentheses(text):
            found = re.findall(r"\((\d+)\)", text)
            if found:
                return int(found[0])
            else:
                alt_found = re.findall(r"\b(\d+)\b", text)
                if alt_found:
                    return int(alt_found[0])
            return None

        fam_match = re.search(r"(\(\d+\))\s+(?:patent )?famil", desc, re.IGNORECASE)
        if fam_match:
            families_count = find_number_in_parentheses(fam_match.group(1))

        pat_match = re.search(r"(\(\d+\))\s+(?:active\s+)?(?:US\s+and\s+international\s+)?(?:patent\s+)?(patents|assets)", desc, re.IGNORECASE)
        if pat_match:
            patents_count = find_number_in_parentheses(pat_match.group(1))

        record = {
            "title": title,
            "description": desc,
            "families_count": families_count,
            "patents_count": patents_count,
            "status": possible_status,
            "run_date": run_date,
            "source_link": source_url,
        }
        debug(
            f"Parsed RZV listing: '{title}' => families={families_count}, "
            f"patents={patents_count}, status={possible_status}", 
            "DEBUG"
        )
        results.append(record)

    debug(f"Total RZV listings parsed: {len(results)}", "INFO")
    return results

def insert_rzv_data(data, conn, session_id):
    table_name = "rzv"
    debug(f"Inserting {len(data)} RZV listings into DB (session_id={session_id})", "INFO")

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
                 families_count, patents_count, status,
                 added_date, newly_added, closed)
                VALUES (%s, %s, %s, %s,
                        %s, %s, %s,
                        %s, %s, %s)
            """, (
                session_id,
                record["source_link"],
                record["title"],
                record["description"],
                record["families_count"],
                record["patents_count"],
                record["status"],
                record["run_date"],
                newly_added,
                closed
            ))
            debug(f"[INSERT] RZV title='{record['title']}' => newly_added={newly_added}, closed={closed}", "DEBUG")

        closed_titles = [t for t in prev_open_titles if t not in current_titles]
        debug(f"Found {len(closed_titles)} newly 'closed' RZV titles in this scrape.", "DEBUG")

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
                SELECT description, families_count, patents_count, status, source_link
                FROM {table_name}
                WHERE session_id = %s AND title = %s AND closed = FALSE
                LIMIT 1
            """, (prev_session_id, ctitle))
            row = cur.fetchone()
            if row:
                desc, fam_ct, pat_ct, st, src = row
                cur.execute(f"""
                    INSERT INTO {table_name}
                    (session_id, source_link, title, description,
                     families_count, patents_count, status,
                     added_date, newly_added, closed)
                    VALUES (%s, %s, %s, %s,
                            %s, %s, %s,
                            NOW(), %s, %s)
                """, (
                    session_id,
                    src,
                    ctitle,
                    desc,
                    fam_ct,
                    pat_ct,
                    st,
                    False,
                    True
                ))
                debug(f"[CLOSED] RZV title='{ctitle}' in session {session_id}", "DEBUG")

        conn.commit()
        debug(f"RZV commit complete for session_id={session_id}.", "INFO")

def run_rzv_scraper():
    debug("Connecting to database for RZV scraping...", "INFO")
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for RZV scraping. Exiting.", "ERROR")
        return

    try:
        table_name = "rzv"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting RZV scrape for session_id={session_id}.", "INFO")

        rzv_url = "https://www.rzv-ip.com/current-patent-opportunities.html"
        html = fetch_page_rzv(rzv_url)
        if not html:
            debug("No HTML content from RZV. Aborting scraper.", "ERROR")
            return

        data = parse_rzv_opportunities(html, rzv_url)
        if not data:
            debug("No RZV data parsed from the page. No insertion performed.", "WARNING")
            return

        insert_rzv_data(data, conn, session_id)

        debug(f"Generating changes-only report for '{table_name}'...", "DEBUG")
        generate_changes_report(conn, table_name)

        debug(f"Finished RZV scraping for session_id={session_id}.", "INFO")
    finally:
        debug("Closing DB connection for RZV scraping...", "DEBUG")
        conn.close()

if __name__ == "__main__":
    run_rzv_scraper()
