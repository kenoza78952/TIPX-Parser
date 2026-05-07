
import time
import re
import psycopg2
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from dbconn import connect_to_db  
from excel_reports import generate_changes_report 

def debug(message, level="INFO"):
    print(f"[{level}] [GTTGRP] {message}")
    
def init_driver():
    debug("Initializing headless Chrome driver.", "DEBUG")
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def parse_fields(text):
    fields = {}
    labels = {
        "TECHNOLOGY AREA(S):": "technology_areas",
        "RELEVANT MARKET(S):": "relevant_markets",
        "EOU:": "eou",
        "JURISDICTION(S):": "jurisdictions"
    }
    for label, key in labels.items():
        pattern = re.compile(re.escape(label) + r'\s*(.*?)\s*(?=<b>|$)', re.DOTALL | re.IGNORECASE)
        match = pattern.search(text)
        if match:
            fields[key] = match.group(1).strip()
            debug(f"Extracted {key}: {fields[key]}", "DEBUG")
        else:
            fields[key] = ""
            debug(f"Field {key} not found.", "DEBUG")
    return fields

def get_listing_data(driver):
    main_url = "https://www.gttgrp.com/patent-acquisition-opportunities/"
    debug(f"Fetching main listing page: {main_url}", "INFO")
    driver.get(main_url)
    
    wait = WebDriverWait(driver, 15)
    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.gdlr-core-blog-grid-content-wrap")))
        debug("Listings have loaded.", "DEBUG")
    except Exception as e:
        debug(f"Error waiting for listings: {e}", "ERROR")
        return []
    
    listings = driver.find_elements(By.CSS_SELECTOR, "div.gdlr-core-blog-grid-content-wrap")
    debug(f"Found {len(listings)} listing elements.", "DEBUG")
    listing_data = []
    for listing in listings:
        try:
            title_element = listing.find_element(By.CSS_SELECTOR, "h3.gdlr-core-blog-title a")
            title = title_element.text.strip()
            detail_url = title_element.get_attribute("href")
            summary = ""
            try:
                summary_element = listing.find_element(By.CSS_SELECTOR, "div.gdlr-core-blog-content")
                summary = summary_element.text.strip()
            except Exception:
                debug("Summary not found for a listing.", "DEBUG")
            debug(f"Extracted listing: {title} | URL: {detail_url}", "INFO")
            listing_data.append({
                "title": title,
                "detail_url": detail_url,
                "summary": summary
            })
        except Exception as e:
            debug(f"Error extracting a listing: {e}", "ERROR")
            continue
    return listing_data

def get_detail_data(driver, url):
    debug(f"Fetching detail page: {url}", "INFO")
    driver.get(url)
    wait = WebDriverWait(driver, 15)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.gdlr-core-text-box-item-content")))
        debug("Detail content loaded.", "DEBUG")
    except Exception as e:
        debug(f"Error waiting for detail content on {url}: {e}", "ERROR")
        return {}
    
    detail_data = {}
    try:
        content_elem = driver.find_element(By.CSS_SELECTOR, "div.gdlr-core-text-box-item-content")
        content_text = content_elem.get_attribute("innerText").strip()
        detail_data["full_text"] = content_text
        debug("Extracted full text from detail page.", "DEBUG")
        
        fields = parse_fields(content_text)
        detail_data.update(fields)
    except Exception as e:
        debug(f"Error extracting detail data from {url}: {e}", "ERROR")
    return detail_data

def insert_gttgrp_data(data, conn, session_id):
    debug("Starting insertion into DB for GTT Group data.", "INFO")
    with conn.cursor() as cur:
        for record in data:
            try:
                cur.execute("SELECT COUNT(*) FROM gttgrp WHERE title = %s", (record["title"],))
                count = cur.fetchone()[0]
                newly_added = (count == 0)
            except Exception as e:
                debug(f"Error checking record {record['title']}: {e}", "ERROR")
                newly_added = True
            closed = False
            try:
                cur.execute("""
                INSERT INTO gttgrp 
                (session_id, title, summary,
                technology_areas, relevant_markets, eou, jurisdictions,
                source_link, newly_added, closed)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                session_id,
                record["title"],
                record["summary"],
                record.get("technology_areas", ""),
                record.get("relevant_markets", ""),
                record.get("eou", ""),
                record.get("jurisdictions", ""),
                record.get("detail_url", ""),  
                newly_added,
                closed
            ))
                debug(f"Inserted record: {record['title']}", "DEBUG")
            except Exception as e:
                debug(f"Error inserting record {record['title']}: {e}", "ERROR")
        conn.commit()
    debug(f"Inserted {len(data)} rows into gttgrp table.", "INFO")

def generate_excel_report(conn):
    try:
        generate_changes_report(conn, "gttgrp")
        debug("Excel report generated.", "INFO")
    except Exception as e:
        debug(f"Error generating Excel report: {e}", "ERROR")

def main():
    driver = init_driver()
    conn = None
    session_id = None
    try:
        conn = connect_to_db()
        with conn.cursor() as cur:
            cur.execute("SELECT COALESCE(MAX(session_id), 0) + 1 FROM gttgrp")
            session_id = cur.fetchone()[0]
        debug(f"New session_id determined: {session_id}", "INFO")
        
        listings = get_listing_data(driver)
        debug(f"Total listings found: {len(listings)}", "INFO")
        if not listings:
            debug("No listings found. Exiting scraper.", "WARNING")
            return
        
        all_data = []
        for idx, listing in enumerate(listings, start=1):
            debug(f"Processing listing {idx}: {listing['title']}", "INFO")
            detail = get_detail_data(driver, listing["detail_url"])
            combined = {**listing, **detail}
            all_data.append(combined)
            time.sleep(1)  # Polite delay
        
        debug(f"Scraping complete. Total listings scraped: {len(all_data)}", "INFO")
        
        insert_gttgrp_data(all_data, conn, session_id)
        
        generate_excel_report(conn)
        
        debug("Scraping and insertion complete. Data has been inserted into DB and report generated.", "INFO")
    except Exception as e:
        debug(f"Exception in main: {e}", "ERROR")
    finally:
        driver.quit()
        if conn:
            conn.close()
        debug("Driver and DB connection closed.", "INFO")

if __name__ == "__main__":
    main()
