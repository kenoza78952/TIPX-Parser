# execute.py

import os
import glob
from datetime import datetime

from debug_utils import debug, save_debug_log, LOG_BUFFER
from vitek import run_vitek_scraper
from transactionip import run_transactionip_scraper
from tangibleip import run_tangible_scraper
from ipinvestments import run_ip_investments_scraper
from dbconn import connect_to_db
from email_utils import send_email_with_attachments

def run_all_scrapers_and_email():

    month_str = datetime.now().strftime("%m") 
    base_month_dir = os.path.join("output", "2025", month_str)

    existing_scrapes = [
        d for d in os.listdir(base_month_dir)
        if os.path.isdir(os.path.join(base_month_dir, d)) and d.startswith("Scrape ")
    ]
    next_scrape_num = len(existing_scrapes) + 1
    timestamp_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    scrape_folder_name = f"Scrape {next_scrape_num} ({timestamp_str})"
    scrape_dir = os.path.join(base_month_dir, scrape_folder_name)
    os.makedirs(scrape_dir, exist_ok=True)

    debug(f"Created subfolder for this scrape: {scrape_dir}", "INFO")

    debug("Starting ALL scrapers...", "INFO")
    run_vitek_scraper()
    run_transactionip_scraper()
    run_tangible_scraper()
    run_ip_investments_scraper()
    debug("All scrapers done.", "INFO")

    generate_full_excel_report(scrape_dir)

    today_str = datetime.now().strftime("%Y%m%d")
    pattern = f"{base_month_dir}/*_changes_report_{today_str}.xlsx"
    xlsx_files = glob.glob(pattern)

    subject = f"Patent Scraper Updates - {datetime.now().strftime('%Y-%m-%d')}"
    if xlsx_files:
        debug(f"Found {len(xlsx_files)} new changes-only report(s): {xlsx_files}", "INFO")
        body = (
            "Hello,\n\n"
            "Please find attached the new patent scraper reports for today.\n"
            "Regards,\n"
            "TIPX Parser"
        )
        send_email_with_attachments(
            to=[""], 
            subject=subject,
            body=body,
            attachments=xlsx_files
        )
    else:
        debug("No new changes-only files found for emailing.", "INFO")
        no_changes_body = (
            "Hello,\n\n"
            "No new changes this week.\n"
            "Regards,\n"
            "TIPX Parser"
        )
        send_email_with_attachments(
            to=[""],
            subject=subject,
            body=no_changes_body,
            attachments=[]
        )

    debug_log_path = os.path.join(scrape_dir, "debug.log")
    save_debug_log(debug_log_path)
    debug(f"Debug log saved to: {debug_log_path}", "INFO")


def generate_full_excel_report(scrape_dir):
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for generating full excel report.", "ERROR")
        return

    table_names = ["vitek", "transactionip", "tangibleip", "ip_investmentsgroup"]
    dataframes = {}
    with conn.cursor() as cur:
        for tname in table_names:
            cur.execute(f"SELECT MAX(session_id) FROM {tname}")
            max_sess = cur.fetchone()[0]
            if not max_sess:
                debug(f"No data in table '{tname}' yet; skipping.", "WARNING")
                continue
            sql = f"SELECT * FROM {tname} WHERE session_id = %s"
            import pandas as pd
            df = pd.read_sql(sql, conn, params=(max_sess,))
            dataframes[tname] = df

    conn.close()

    import pandas as pd
    full_report_path = os.path.join(scrape_dir, "Scrape_Report.xlsx")
    with pd.ExcelWriter(full_report_path) as writer:
        for tname, df in dataframes.items():
            df.to_excel(writer, sheet_name=tname, index=False)

    debug(f"Full scrape data saved to: {full_report_path}", "INFO")


def run_all_scrapers_test_mode():
    debug("[TEST] Starting scrapers in TEST MODE. Email = suppressed or test-only.", "INFO")
    
    run_vitek_scraper()
    run_transactionip_scraper()
    run_tangible_scraper()
    run_ip_investments_scraper()
    
    debug("[TEST] Completed scrapers in TEST MODE. Check logs for details.", "INFO")


if __name__ == "__main__":
    run_all_scrapers_and_email()
