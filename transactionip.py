# transactionip.py

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

def fetch_page_transactionip(url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200:
            return resp.text
        else:
            debug(f"Failed to fetch page: {url}, status {resp.status_code}", "ERROR")
            return None
    except requests.exceptions.RequestException as e:
        debug(f"Error fetching page {url}: {e}", "ERROR")
        return None

def parse_main_page(html_content):
    if not html_content:
        debug("No HTML content to parse in parse_main_page()", "ERROR")
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    wrapper_divs = soup.find_all("div", class_="wpb_wrapper")
    categories = []

    if len(wrapper_divs) > 1:
        p_tag = wrapper_divs[1].find("p", align="center")
        if p_tag:
            a_tags = p_tag.find_all("a", href=True)
            for a_tag in a_tags:
                if a_tag.find("strong"):
                    cat_name = a_tag.find("strong").text.strip()
                elif a_tag.parent and a_tag.parent.name == "strong":
                    cat_name = a_tag.text.strip()
                else:
                    cat_name = "Unknown Category"

                cat_link = a_tag["href"]
                subdata = parse_category_page(cat_link)
                categories.append({
                    "Category Name": cat_name,
                    "Category Link": cat_link,
                    "Page Data": subdata
                })
    return categories

def parse_category_page(url):
    html = fetch_page_transactionip(url)
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    post_divs = soup.find_all("div", class_="elementor-post__text")
    data = []

    for div in post_divs:
        h3_tag = div.find("h3", class_="elementor-post__title")
        if h3_tag and h3_tag.find("a"):
            link_tag = h3_tag.find("a")
            title = link_tag.text.strip()
            detail_link = link_tag.get("href", url)
        else:
            title = "Unknown Title"
            detail_link = url

        p_elements = div.find_all("p")
        short_desc = [p.text.strip() for p in p_elements if p.text.strip()]

        full_desc = fetch_and_parse_detail_page(detail_link)
        descriptions = full_desc if full_desc else short_desc

        descriptions = clean_descriptions(descriptions)

        data.append({
            "Title": title,
            "Descriptions": descriptions,
            "SourceLink": detail_link
        })
    return data

def fetch_and_parse_detail_page(detail_url):
    detail_html = fetch_page_transactionip(detail_url)
    if not detail_html:
        return []
    soup = BeautifulSoup(detail_html, "html.parser")
    paragraphs = soup.find_all("p")
    return [p.text.strip() for p in paragraphs if p.text.strip()]

def clean_descriptions(desc_list):
    cleaned = []
    for text in desc_list:
        if not text:
            continue
        text = text.replace("\xa0", " ")
        text = text.replace("[email protected]", "[email protected]")
        cleaned.append(text.strip())
    return cleaned

def insert_patents_transactionip(data, conn, session_id):
    table_name = "transactionip"
    with conn.cursor() as cur:
        latest_titles = [d["Title"] for d in data]
        prev_session_id = session_id - 1

        cur.execute(
            f"SELECT title FROM {table_name} WHERE session_id = %s AND closed = FALSE",
            (prev_session_id,)
        )
        prev_titles = [row[0] for row in cur.fetchall()]

        for entry in data:
            cur.execute(f"""
                SELECT COUNT(*) 
                FROM {table_name}
                WHERE session_id = %s
                  AND title = %s
            """, (session_id, entry["Title"]))
            already_inserted = cur.fetchone()[0]
            if already_inserted > 0:
                continue

            cur.execute(
                f"SELECT COUNT(*) FROM {table_name} WHERE title = %s",
                (entry["Title"],)
            )
            existed = (cur.fetchone()[0] > 0)
            newly_added = not existed
            closed = False

            cur.execute(f"""
                INSERT INTO {table_name}
                (session_id, category, category_link,
                 title, description, newly_added, closed,
                 source_link, added_date)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
            """, (
                session_id,
                entry.get("Category Name", "Unknown"),
                entry.get("Category Link", ""),
                entry["Title"],
                str(entry["Descriptions"]),
                newly_added,
                closed,
                entry.get("SourceLink", "")
            ))
            debug(f"[INSERT] '{entry['Title']}' (newly_added={newly_added}, closed={closed})", "DEBUG")

        closed_titles = [t for t in prev_titles if t not in latest_titles]
        for ctitle in closed_titles:
            cur.execute(f"""
                SELECT COUNT(*)
                FROM {table_name}
                WHERE session_id = %s
                  AND title = %s
                  AND closed = TRUE
            """, (session_id, ctitle))
            already_closed = cur.fetchone()[0]
            if already_closed > 0:
                continue

            cur.execute(f"""
                SELECT category, category_link, description, source_link
                FROM {table_name}
                WHERE session_id = %s
                  AND title = %s
                  AND closed = FALSE
                LIMIT 1
            """, (prev_session_id, ctitle))
            row = cur.fetchone()
            if row:
                category, cat_link, desc, src = row
                cur.execute(f"""
                    INSERT INTO {table_name}
                    (session_id, category, category_link,
                     title, description, newly_added, closed,
                     source_link, added_date)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
                """, (
                    session_id,
                    category,
                    cat_link,
                    ctitle,
                    desc,
                    False,
                    True,
                    src
                ))
                debug(f"[CLOSED] Title='{ctitle}' in session {session_id}", "DEBUG")

        conn.commit()
        debug(f"TransactionIP commit complete for session_id={session_id}.", "DEBUG")

def run_transactionip_scraper():
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for TransactionsIP", "ERROR")
        return

    try:
        table_name = "transactionip"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting TransactionsIP scrape for session_id={session_id}.", "INFO")

        main_url = "https://transactionsip.com/patents-for-sale/"
        main_page = fetch_page_transactionip(main_url)
        if not main_page:
            debug("Could not fetch main TransactionsIP page.", "ERROR")
            return

        cats_data = parse_main_page(main_page)
        all_listings = []
        run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for cat_obj in cats_data:
            c_name = cat_obj["Category Name"]
            c_link = cat_obj["Category Link"]
            for patent_info in cat_obj["Page Data"]:
                all_listings.append({
                    "Category Name": c_name,
                    "Category Link": c_link,
                    "Title": patent_info["Title"],
                    "Descriptions": patent_info["Descriptions"],
                    "Run Date": run_date,
                    "SourceLink": patent_info.get("SourceLink", c_link)
                })

        if not all_listings:
            debug("No sub-listings found in TransactionsIP.", "WARNING")
            return

        insert_patents_transactionip(all_listings, conn, session_id)

        generate_changes_report(conn, table_name)

        debug(f"TransactionsIP scraping complete for session_id={session_id}", "INFO")

    finally:
        conn.close()

if __name__ == "__main__":
    run_transactionip_scraper()
