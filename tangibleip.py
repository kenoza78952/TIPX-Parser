# tangibleip.py

import requests
from datetime import datetime
from bs4 import BeautifulSoup

from dbconn import connect_to_db
from debug_utils import debug
from excel_reports import generate_changes_report

def get_new_session_id(conn, table_name):
    with conn.cursor() as cur:
        cur.execute(f"SELECT COALESCE(MAX(session_id), 0) + 1 FROM {table_name}")
        return cur.fetchone()[0]

def fetch_page_tangible(url):
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            return resp.text
        else:
            debug(f"Failed to fetch Tangible IP page: {resp.status_code}", "ERROR")
            return None
    except Exception as e:
        debug(f"Error fetching Tangible IP: {e}", "ERROR")
        return None

def parse_patents_tangible(html_content):
    if not html_content:
        debug("No HTML content to parse for TangibleIP.", "ERROR")
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    portfolio_items = soup.find_all("div", class_="et_pb_portfolio_item")
    debug(f"Found {len(portfolio_items)} portfolio items in TangibleIP HTML.", "DEBUG")

    data = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for item in portfolio_items:
        link_tag = item.find("a")
        if not link_tag:
            continue

        patent_url = link_tag["href"]
        patent_html = fetch_page_tangible(patent_url)
        if not patent_html:
            debug(f"Could not fetch detail page: {patent_url}", "WARNING")
            continue

        patent_soup = BeautifulSoup(patent_html, "html.parser")
        text_sections = patent_soup.find_all("div", class_="et_pb_text_inner")

        if len(text_sections) >= 4:
            title = text_sections[0].text.strip()
            patent_id = text_sections[1].text.strip()
            summary = text_sections[2].text.strip()
            description = text_sections[3].text.strip()
            data.append({
                "Title": title,
                "ID": patent_id,
                "Summary": summary,
                "Description": description,
                "Run Date": run_date,
                "SourceLink": patent_url
            })
            debug(f"Parsed TangibleIP listing: {title} (ID={patent_id})", "DEBUG")

    return data

def insert_patents_tangible(data, conn, session_id):
    table_name = "tangibleip"
    with conn.cursor() as cur:
        current_ids = [d["ID"] for d in data]

        prev_session_id = session_id - 1
        cur.execute(f"""
            SELECT tangibleip_id, closed
            FROM {table_name}
            WHERE session_id = %s
        """, (prev_session_id,))
        prev_rows = cur.fetchall()

        prev_open_ids = {r[0] for r in prev_rows if r[1] == False}

        for record in data:
            cur.execute("SELECT COUNT(*) FROM tangibleip WHERE tangibleip_id = %s", (record["ID"],))
            existed_before = (cur.fetchone()[0] > 0)
            newly_added = not existed_before

            cur.execute(f"""
                INSERT INTO {table_name}
                  (title, session_id, tangibleip_id,
                   summary, description, added_date,
                   newly_added, closed, source_link)
                VALUES (%s, %s, %s, %s, %s, NOW(), %s, %s, %s)
            """, (
                record["Title"],
                session_id,
                record["ID"],
                record["Summary"],
                record["Description"],
                newly_added,
                False,
                record["SourceLink"]
            ))
            debug(f"[INSERT] TangibleIP ID={record['ID']} newly_added={newly_added}", "DEBUG")

        newly_closed = [pid for pid in prev_open_ids if pid not in current_ids]
        for cid in newly_closed:
            cur.execute(f"""
                SELECT title, summary, description, source_link
                FROM {table_name}
                WHERE tangibleip_id = %s AND session_id = %s
            """, (cid, prev_session_id))
            old_data = cur.fetchone()
            if old_data:
                title, summary, description, src_link = old_data
                cur.execute(f"""
                    INSERT INTO {table_name}
                      (title, session_id, tangibleip_id,
                       summary, description, added_date,
                       newly_added, closed, source_link)
                    VALUES (%s, %s, %s, %s, %s, NOW(), %s, %s, %s)
                """, (
                    title,
                    session_id,
                    cid,
                    summary,
                    description,
                    False,
                    True,
                    src_link
                ))
                debug(f"[CLOSED] TangibleIP ID={cid} closed in session {session_id}", "DEBUG")

    conn.commit()
    debug(f"TangibleIP commit complete for session_id={session_id}.", "DEBUG")

def run_tangible_scraper():
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for TangibleIP", "ERROR")
        return

    try:
        table_name = "tangibleip"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting TangibleIP scrape for session_id={session_id}.", "INFO")

        tangible_url = "https://tangibleip.biz/patents-for-sale"
        html = fetch_page_tangible(tangible_url)
        if not html:
            debug("Could not fetch Tangible IP page.", "ERROR")
            return

        data = parse_patents_tangible(html)
        if not data:
            debug("No data parsed from Tangible IP details.", "WARNING")
            return

        insert_patents_tangible(data, conn, session_id)
        generate_changes_report(conn, table_name)

        debug("All done with TangibleIP scraping.", "INFO")
    finally:
        conn.close()

if __name__ == "__main__":
    run_tangible_scraper()
