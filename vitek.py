# vitek.py
import time
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from dbconn import connect_to_db
from debug_utils import debug
from excel_reports import generate_changes_report

def get_new_session_id(conn, table_name):
    with conn.cursor() as cur:
        cur.execute(f"SELECT COALESCE(MAX(session_id), 0) + 1 FROM {table_name}")
        return cur.fetchone()[0]

def fetch_page_vitek(driver, url):
    debug(f"Navigating to: {url}", "DEBUG")
    driver.get(url)
    time.sleep(2)
    page_source = driver.page_source
    debug("Page source fetched successfully.", "DEBUG")
    return page_source

def parse_patents_vitek(html_content, source_url):
    soup = BeautifulSoup(html_content, 'html.parser')
    all_articles = soup.find_all('div', class_='wp-block-media-text__content')

    data = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    debug(f"Found {len(all_articles)} content blocks on the Vitek page.", "DEBUG")

    for article in all_articles:
        title_tag = article.find('h2')
        if not title_tag or not title_tag.text.strip():
            debug("Skipping an article with no valid <h2> title.", "DEBUG")
            continue

        raw_title = title_tag.text.strip()
        p_elements = article.find_all('p')

        description = p_elements[0].text.strip() if len(p_elements) > 0 else ""
        summary = p_elements[1].text.strip() if len(p_elements) > 1 else ""
        timeline = p_elements[2].text.strip() if len(p_elements) > 2 else ""

        if "Transaction anticipated" not in timeline and "expected" not in timeline:
            timeline = ""

        data.append({
            "Title": raw_title,
            "Description": description,
            "Summary": summary,
            "Timeline": timeline,
            "SourceLink": source_url,
            "Run Date": run_date
        })
        debug(f"Parsed item: '{raw_title}'", "DEBUG")

    return data

def insert_patents_vitek(data, conn, session_id):
    table_name = 'vitek'
    current_titles = [d["Title"] for d in data]

    with conn.cursor() as cur:
        prev_session_id = session_id - 1
        cur.execute(
            f"SELECT title FROM {table_name} WHERE session_id = %s AND closed = FALSE",
            (prev_session_id,)
        )
        prev_open_titles = [r[0] for r in cur.fetchall()]
        debug(f"Session {prev_session_id} had {len(prev_open_titles)} open titles.", "DEBUG")

        for record in data:
            title = record["Title"]
            cur.execute(f"SELECT COUNT(*) FROM {table_name} WHERE title = %s", (title,))
            seen_before = (cur.fetchone()[0] > 0)
            newly_added = not seen_before
            closed = False

            cur.execute(f"""
                INSERT INTO {table_name}
                (session_id, title, description, summary, timeline, source_link,
                 newly_added, closed, added_date)
                VALUES (%s, %s, %s, %s, %s, %s,
                        %s, %s, NOW())
            """, (
                session_id,
                title,
                record["Description"],
                record["Summary"],
                record["Timeline"],
                record["SourceLink"],
                newly_added,
                closed
            ))
            debug(f"[INSERT] '{title}' => newly_added={newly_added}, closed={closed}", "DEBUG")

        disappeared = [t for t in prev_open_titles if t not in current_titles]
        debug(f"{len(disappeared)} titles disappeared from the previous session.", "DEBUG")

        for ctitle in disappeared:
            cur.execute(f"""
                INSERT INTO {table_name}
                (session_id, title, description, summary, timeline, source_link,
                 newly_added, closed, added_date)
                SELECT
                  %s, title, description, summary, timeline, source_link,
                  FALSE, TRUE, NOW()
                FROM {table_name}
                WHERE session_id = %s AND title = %s AND closed = FALSE
                LIMIT 1
            """, (session_id, prev_session_id, ctitle))
            debug(f"[CLOSED] Title='{ctitle}' in session {session_id}", "DEBUG")

    conn.commit()
    debug(f"Vitek commit complete for session_id={session_id}.", "DEBUG")

def run_vitek_scraper():
    debug("Connecting to database...", "INFO")
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB. Aborting Vitek scraper.", "ERROR")
        return
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)

    try:
        table_name = 'vitek'
        session_id = get_new_session_id(conn, table_name)
        vitek_url = 'https://vitek-ip.com/portfolios-for-sale/'

        debug(f"Starting Vitek scrape for session_id={session_id}.", "INFO")
        page_source = fetch_page_vitek(driver, vitek_url)

        data = parse_patents_vitek(page_source, vitek_url)
        debug(f"Parsed {len(data)} items from Vitek page.", "INFO")

        insert_patents_vitek(data, conn, session_id)

        generate_changes_report(conn, table_name)

        debug(f"Completed Vitek scraping for session_id={session_id}.", "INFO")
    finally:
        driver.quit()
        conn.close()

if __name__ == "__main__":
    run_vitek_scraper()
