# ip_investmentsgroup.py

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

def fetch_page_investments(url, retries=3, backoff_factor=0.3):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )
    }
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.text
            else:
                debug(f"Failed to fetch page: {url}, status {response.status_code}", "ERROR")
                return None
        except requests.exceptions.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            debug(f"Error fetching page: {e}. Retrying in {wait_time} secs...", "WARNING")
            time.sleep(wait_time)

    debug("Max retries reached. Failed to fetch the page.", "ERROR")
    return None

def parse_patents_investments(html_content, base_url):
    if not html_content:
        debug("No content to parse for IP Investments Group.", "ERROR")
        return []

    soup = BeautifulSoup(html_content, "html.parser")
    all_articles = soup.find_all(class_="offeringBlock")
    data = []
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for article in all_articles:
        title_div = article.find("div", style="float:left;")
        if not title_div:
            continue
        h2_tag = title_div.find("h2")
        link_tag = h2_tag.find("a") if h2_tag else None

        if link_tag:
            offering_title = link_tag.text.strip()
            source_link = link_tag.get("href", base_url)
        else:
            offering_title = "No Title Found"
            source_link = base_url
        description = None
        p_tags = article.find_all("p")
        for p in p_tags:
            text = p.get_text(strip=True)
            if text:
                description = text
                break

        patents_list = article.find("ul", class_="offering-fields")
        if not patents_list:
            data.append({
                "Title": offering_title,
                "Description": description,
                "PatentNumbers": "",
                "PatentTitles": "",
                "Run Date": run_date,
                "SourceLink": source_link
            })
            continue

        valid_li = []
        all_li = patents_list.find_all("li")
        for li in all_li:
            raw_text = li.get_text(strip=True)
            lower_check = raw_text.lower().replace(".", "")
            if lower_check in ["patent no", "title"]:
                continue
            valid_li.append(raw_text)

        pair_set = set()
        for i in range(0, len(valid_li), 2):
            if i + 1 < len(valid_li):
                raw_number = valid_li[i].replace(",", "").replace("\xa0", " ").strip()
                raw_title = valid_li[i + 1].replace("\xa0", " ")
                try:
                    num_as_int = int(raw_number)
                    clean_number = str(num_as_int)
                except ValueError:
                    clean_number = re.sub(r"\s+", " ", raw_number)

                lower_title = raw_title.lower().strip()
                clean_title = re.sub(r"\s+", " ", lower_title)
                pair_set.add((clean_number, clean_title))

        patent_number_list = []
        patent_title_list = []
        for (num_str, title_str) in pair_set:
            patent_number_list.append(num_str)
            patent_title_list.append(title_str)

        patent_number_list.sort()
        patent_title_list.sort()

        for idx, t in enumerate(patent_title_list):
            patent_title_list[idx] = f"- {t}"

        patent_numbers_str = "\n".join(patent_number_list)
        patent_titles_str = "\n".join(patent_title_list)

        data.append({
            "Title": offering_title,
            "Description": description,
            "PatentNumbers": patent_numbers_str,
            "PatentTitles": patent_titles_str,
            "Run Date": run_date,
            "SourceLink": source_link
        })
        debug(f"Parsed IP Investments offering: '{offering_title}'", "DEBUG")

    return data

def insert_patents_investments(data, conn, session_id):
    table_name = "ip_investmentsgroup"
    with conn.cursor() as cur:
        latest_titles = [d["Title"] for d in data]
        prev_session_id = session_id - 1

        cur.execute(f"""
            SELECT title FROM {table_name}
            WHERE session_id = %s AND closed = FALSE
        """, (prev_session_id,))
        prev_titles = [row[0] for row in cur.fetchall()]

        for entry in data:
            cur.execute("SELECT COUNT(*) FROM ip_investmentsgroup WHERE title = %s", (entry["Title"],))
            exists = cur.fetchone()[0]
            newly_added = (exists == 0)
            closed = False

            cur.execute(f"""
                INSERT INTO {table_name}
                (title, session_id, description, patent_numbers, patent_titles,
                 added_date, newly_added, closed, source_link)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                entry["Title"],
                session_id,
                entry["Description"],
                entry["PatentNumbers"],
                entry["PatentTitles"],
                entry["Run Date"],
                newly_added,
                closed,
                entry["SourceLink"]
            ))
            debug(f"[INSERT] '{entry['Title']}' (newly_added={newly_added}, closed={closed})", "DEBUG")

        closed_titles = [t for t in prev_titles if t not in latest_titles]
        for ctitle in closed_titles:
            cur.execute(f"""
                SELECT description, patent_numbers, patent_titles, source_link
                FROM {table_name}
                WHERE title = %s AND session_id = %s AND closed = FALSE
            """, (ctitle, prev_session_id))
            row = cur.fetchone()
            if row:
                description, patent_numbers, patent_titles, src_link = row
                cur.execute(f"""
                    INSERT INTO {table_name}
                    (title, session_id, description, patent_numbers, patent_titles,
                     added_date, newly_added, closed, source_link)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    ctitle,
                    session_id,
                    description,
                    patent_numbers,
                    patent_titles,
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    False,
                    True,
                    src_link
                ))
                debug(f"[CLOSED] '{ctitle}' in session_id={session_id}", "DEBUG")

        conn.commit()
        debug(f"IP Investments commit complete for session_id={session_id}.", "DEBUG")

def run_ip_investments_scraper():
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to the database for IP Investments Group.", "ERROR")
        return

    try:
        table_name = "ip_investmentsgroup"
        session_id = get_new_session_id(conn, table_name)
        debug(f"Starting IP Investments Group scrape for session_id={session_id}.", "INFO")

        investments_url = "https://ipinvestmentsgroup.com/current-offerings/"
        html = fetch_page_investments(investments_url)
        if html:
            data = parse_patents_investments(html, base_url=investments_url)
            insert_patents_investments(data, conn, session_id)
            generate_changes_report(conn, table_name)
            debug(f"Finished IP Investments scraping for session_id={session_id}.", "INFO")
        else:
            debug("Could not retrieve IP Investments page content.", "ERROR")
    finally:
        conn.close()

if __name__ == "__main__":
    run_ip_investments_scraper()
