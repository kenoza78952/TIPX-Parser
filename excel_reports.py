# excel_reports.py
import os
import sys
import shutil
import pandas as pd
from datetime import datetime
from debug_utils import debug
from dbconn import connect_to_db

YEAR_DIR = f"output/{YEAR}"
os.makedirs(YEAR_DIR, exist_ok=True)

# Make 12 month dirs
for m in range(1, 13):
    os.makedirs(os.path.join(YEAR_DIR, f"{m:02d}"), exist_ok=True)

FULL_DB_DIR = os.path.join(YEAR_DIR, "FullDB")
os.makedirs(FULL_DB_DIR, exist_ok=True)

def bool_to_yes_no(val):
    if val is True:
        return "Yes"
    elif val is False:
        return "No"
    return val 

def generate_changes_report(conn, table_name):
    with conn.cursor() as cur:
        cur.execute(f"SELECT MAX(session_id) FROM {table_name}")
        max_session_id = cur.fetchone()[0]

    if not max_session_id:
        debug(f"No rows found in table='{table_name}'. Nothing to report.", "DEBUG")
        return

    query_new = f"SELECT * FROM {table_name} WHERE session_id = %s AND newly_added = TRUE"
    query_closed = f"SELECT * FROM {table_name} WHERE session_id = %s AND closed = TRUE"

    if table_name == "parallelnorth":
        query_status_changed = f"SELECT * FROM {table_name} WHERE session_id = %s AND status_changed = TRUE"
    else:
        query_status_changed = None
        
    df_new = pd.read_sql(query_new, conn, params=(max_session_id,))
    df_closed = pd.read_sql(query_closed, conn, params=(max_session_id,))
    if query_status_changed:
        df_status_changed = pd.read_sql(query_status_changed, conn, params=(max_session_id,))
    else:
        df_status_changed = pd.DataFrame()

    debug(f"Table='{table_name}', session={max_session_id}, new={len(df_new)}, closed={len(df_closed)}", "DEBUG")

    # if empty => skip
    if df_new.empty and df_closed.empty and df_status_changed.empty:
        debug("No new or closed or status_changed items => skipping Excel generation.", "DEBUG")
        return

    month_str = datetime.now().strftime("%m")
    month_dir = os.path.join(YEAR_DIR, month_str)
    os.makedirs(month_dir, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    filename = os.path.join(month_dir, f"{table_name}_changes_report_{date_str}.xlsx")

    df_new = rename_and_reorder_dataframe(df_new, table_name)
    df_closed = rename_and_reorder_dataframe(df_closed, table_name)
    df_status_changed = rename_and_reorder_dataframe(df_status_changed, table_name)

    with pd.ExcelWriter(filename) as writer:
        if not df_new.empty:
            df_new.to_excel(writer, sheet_name="new_records", index=False)
        if not df_closed.empty:
            df_closed.to_excel(writer, sheet_name="closed_records", index=False)
        if not df_status_changed.empty:
            df_status_changed.to_excel(writer, sheet_name="status_changed", index=False)

    debug(f"Changes-only report generated: {filename}", "INFO")

def rename_and_reorder_dataframe(df, table_name):
    if table_name == "ip_investmentsgroup":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "title": "Title",
            "description": "Description",
            "patent_numbers": "Patent Number",
            "patent_titles": "Patent Title",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID", "Session ID", "Link", "Title", "Description",
            "Patent Number", "Patent Title", "Date Added",
            "New Listing (Yes/No)", "Closed (Yes/No)"
        ]
    elif table_name == "tangibleip":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "title": "Title",
            "tangibleip_id": "TANIP ID",
            "summary": "Summary",
            "description": "Description",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Link","Title","TANIP ID","Summary",
            "Description","Date Added","New Listing (Yes/No)","Closed (Yes/No)"
        ]
    elif table_name == "transactionip":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "category": "Category",
            "category_link": "Category Link",
            "title": "Title",
            "description": "Description",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Link","Category","Category Link","Title",
            "Description","Date Added","New Listing (Yes/No)","Closed (Yes/No)"
        ]
    elif table_name == "vitek":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "title": "Title",
            "description": "Description",
            "summary": "Summary",
            "timeline": "Timeline",
            "source_link": "Link",
            "sold": "Sold (Yes/No)",
            "licensed": "Licensed (Yes/No)",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Title","Description","Summary","Timeline","Link",
            "Sold (Yes/No)","Licensed (Yes/No)","Date Added",
            "New Listing (Yes/No)","Closed (Yes/No)"
        ]
    elif table_name == "rzv":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "title": "Title",
            "description": "Description",
            "families_count": "Families",
            "patents_count": "Patents",
            "status": "Status",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Link","Title","Description",
            "Families","Patents","Status","Date Added",
            "New Listing (Yes/No)","Closed (Yes/No)"
        ]
    elif table_name == "icap":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "title": "Title",
            "description": "Description",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Link","Title","Description","Date Added",
            "New Listing (Yes/No)","Closed (Yes/No)"
        ]
    elif table_name == "parallelnorth":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",
            "sector": "Sector",
            "seller": "Seller",
            "sale_type": "Sale/License Type",
            "status": "Status",
            "patent_count": "# Patents",
            "status_changed": "Status Changed (Yes/No)",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID", "Session ID", "Link", "Sector", "Seller", "Sale/License Type",
            "Status", "# Patents", "Date Added", 
            "New Listing (Yes/No)", "Closed (Yes/No)", "Status Changed (Yes/No)"
        ]
    elif table_name == "ipapproach":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "source_link": "Link",  
            "category": "Category",
            "title": "Title",
            "description": "Description",
            "patent_numbers": "Patent Numbers",
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID","Session ID","Link","Category","Title","Description",
            "Patent Numbers","Date Added",
            "New Listing (Yes/No)","Closed (Yes/No)"
        ]

    elif table_name == "gttgrp":
        col_map = {
            "id": "ID",
            "session_id": "Session ID",
            "title": "Title",
            "summary": "Summary",
            "technology_areas": "Technology Area(s)",
            "relevant_markets": "Relevant Market(s)",
            "eou": "EOU",
            "jurisdictions": "Jurisdiction(s)",
            "source_link": "Link",
            "detail_url": "Detail URL",   
            "full_text": "Full Text",     
            "added_date": "Date Added",
            "newly_added": "New Listing (Yes/No)",
            "closed": "Closed (Yes/No)",
        }
        final_order = [
            "ID", "Session ID", "Title", "Summary",
            "Technology Area(s)", "Relevant Market(s)", "EOU", "Jurisdiction(s)",
            "Link", 
            "Detail URL",      
            "Full Text",       
            "Date Added", 
            "New Listing (Yes/No)", 
            "Closed (Yes/No)"
        ]
    else:
        col_map = {}
        final_order = list(df.columns)
        
    df = df.rename(columns=col_map)

    for col in df.columns:
        if col.endswith("(Yes/No)"):
            df[col] = df[col].apply(bool_to_yes_no)

    existing = [c for c in final_order if c in df.columns]
    leftover = [c for c in df.columns if c not in existing]
    return df.reindex(columns=existing + leftover)

def clear_directory(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path, exist_ok=True)
        debug(f"Directory created: {dir_path}", "INFO")
        return
    for entry in os.scandir(dir_path):
        if entry.is_file() or entry.is_symlink():
            os.remove(entry.path)
        elif entry.is_dir():
            shutil.rmtree(entry.path)
    debug(f"Cleared contents of directory: {dir_path}", "INFO")

def delete_year_folder(year_str):
    year_path = os.path.join("output", year_str)
    if os.path.exists(year_path):
        shutil.rmtree(year_path, ignore_errors=True)
        debug(f"Deleted year folder: {year_path}", "INFO")
    else:
        debug(f"Year folder '{year_path}' does not exist.", "WARNING")

def delete_multiple_months(year_str, months_list):
    for m in months_list:
        month_path = os.path.join("output", year_str, m)
        if os.path.exists(month_path):
            shutil.rmtree(month_path, ignore_errors=True)
            debug(f"Deleted month folder: {month_path}", "INFO")
        else:
            debug(f"Month folder '{month_path}' does not exist.", "WARNING")

def delete_scrape_number(year_str, month_str, scrape_number):
    base_dir = os.path.join("output", year_str, month_str)
    if not os.path.isdir(base_dir):
        debug(f"Month directory '{base_dir}' does not exist.", "WARNING")
        return

    subfolders = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d))]
    found = False
    for folder in subfolders:
        if folder.startswith(f"Scrape {scrape_number} ("):
            target_path = os.path.join(base_dir, folder)
            shutil.rmtree(target_path, ignore_errors=True)
            debug(f"Deleted scrape folder: {target_path}", "INFO")
            found = True
    if not found:
        debug(f"No scrape folder 'Scrape {scrape_number} (...)' found in {base_dir}.", "INFO")

def export_full_db():
    conn = connect_to_db()
    if not conn:
        debug("Could not connect to DB for Full DB export.", "ERROR")
        return

    table_names = ["vitek", "transactionip", "tangibleip", "ip_investmentsgroup"]
    dataframes = {}

    try:
        with conn.cursor() as cur:
            for tname in table_names:
                sql = f"SELECT * FROM {tname}"
                # same read_sql warning may appear
                df = pd.read_sql(sql, conn)
                dataframes[tname] = df
    except Exception as e:
        debug(f"Error exporting full DB: {e}", "ERROR")
    finally:
        conn.close()

    date_str = datetime.now().strftime("%Y%m%d")
    full_db_path = os.path.join(FULL_DB_DIR, f"FullDB_{date_str}.xlsx")

    with pd.ExcelWriter(full_db_path) as writer:
        for tname, df in dataframes.items():
            df.to_excel(writer, sheet_name=tname, index=False)

    debug(f"Full DB exported => {full_db_path}", "INFO")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        if cmd == "export_full_db":
            export_full_db()
        else:
            print("Unknown command. e.g. python excel_reports.py export_full_db")
    else:
        print("Usage: python excel_reports.py export_full_db")
