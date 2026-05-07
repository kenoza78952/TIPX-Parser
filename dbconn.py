# dbcon.py
import os
import sys
import psycopg2
import sqlalchemy
import pandas as pd

DB_NAME = os.getenv("DB_NAME")

def connect_to_db():
    try:
        conn = psycopg2.connect(
            dbname=DB_NAME,
            user=DB_USER,
            password=DB_PASSWORD,
            host=DB_HOST,
            port=DB_PORT
        )
        create_tables(conn)
        print(f"✅ Connected to TIPX server and '{DB_NAME}' database.")
        return conn
    except Exception as e:
        print(f"❌ Error connecting to TIPX server: {e}")
        return None

def create_tables(conn):
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS ip_investmentsgroup (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    title TEXT,
                    description TEXT,
                    patent_numbers TEXT,
                    patent_titles TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS tangibleip (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    title TEXT,
                    tangibleip_id TEXT,
                    summary TEXT,
                    description TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS transactionip (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    category TEXT,
                    category_link TEXT,
                    title TEXT,
                    description TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS vitek (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    title TEXT,
                    description TEXT,
                    summary TEXT,
                    timeline TEXT,
                    source_link TEXT,
                    sold BOOLEAN DEFAULT FALSE,
                    licensed BOOLEAN DEFAULT FALSE,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS rzv (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    title TEXT,
                    description TEXT,
                    families_count INT,
                    patents_count INT,
                    status TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS icap (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    title TEXT,
                    description TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS parallelnorth (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,
                    sector TEXT,
                    seller TEXT,
                    sale_type TEXT,
                    status TEXT,  -- "Open" or "Closed"
                    patent_count TEXT,  
                    status_changed BOOLEAN DEFAULT FALSE,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS ipapproach (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    source_link TEXT,    
                    category TEXT,
                    title TEXT,
                    description TEXT,
                    patent_numbers TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS gttgrp (
                    id SERIAL PRIMARY KEY,
                    session_id INT,
                    title TEXT,
                    summary TEXT,
                    technology_areas TEXT,
                    relevant_markets TEXT,
                    eou TEXT,
                    jurisdictions TEXT,
                    source_link TEXT,
                    added_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    newly_added BOOLEAN,
                    closed BOOLEAN DEFAULT FALSE
                );
            """)

            conn.commit()
            print("✅ Tables confirmed: 'ip_investmentsgroup', 'tangibleip', 'transactionip', 'vitek', 'rzv', 'icap'.")
    except Exception as e:
        print(f"❌ Error creating/verifying tables: {e}")

def get_sqlalchemy_engine():
    return sqlalchemy.create_engine(
        f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    )


def import_excel_to_table(table_name, excel_path):
    table_column_map = {
        "ip_investmentsgroup": {
            "Session ID": "session_id",
            "Link": "source_link",
            "Title": "title",
            "Description": "description",
            "Patent Number": "patent_numbers",
            "Patent Title": "patent_titles",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
        "tangibleip": {
            "Session ID": "session_id",
            "Link": "source_link",
            "Title": "title",
            "TANIP ID": "tangibleip_id",
            "Summary": "summary",
            "Description": "description",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
        "transactionip": {
            "Session ID": "session_id",
            "Link": "source_link",
            "Category": "category",
            "Category Link": "category_link",
            "Title": "title",
            "Description": "description",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
        "vitek": {
            "Session ID": "session_id",
            "Title": "title",
            "Description": "description",
            "Summary": "summary",
            "Timeline": "timeline",
            "Link": "source_link",
            "Sold (Yes/No)": "sold",
            "Licensed (Yes/No)": "licensed",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
            "rzv": {
            "Session ID": "session_id",
            "Link": "source_link",
            "Title": "title",
            "Description": "description",
            "Families Count": "families_count",
            "Patents Count": "patents_count",
            "Status": "status",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
        "icap": {
            "Session ID": "session_id",
            "Link": "source_link",
            "Title": "title",
            "Description": "description",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
        "ipapproach": {
            "Session ID": "session_id",
            "Source Link": "source_link",
            "Category": "category",
            "Title": "title",
            "Description": "description",
            "Patent Numbers": "patent_numbers",
            "Date Added": "added_date",
            "New Listing (Yes/No)": "newly_added",
            "Closed (Yes/No)": "closed",
        },
    }
    
    if table_name not in table_column_map:
        print(f"❌ Unknown table_name='{table_name}'. Must be one of: ip_investmentsgroup, tangibleip, transactionip, vitek.")
        return

    col_map = table_column_map[table_name]

    # Read Excel
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"❌ Error reading Excel file '{excel_path}': {e}")
        return

    if df.empty:
        print(f"⚠️ The Excel file '{excel_path}' is empty. Nothing to insert.")
        return
    rename_dict = {excel_col: db_col for excel_col, db_col in col_map.items() if excel_col in df.columns}
    df = df.rename(columns=rename_dict)

    yes_no_cols = [db_col for db_col in rename_dict.values() if db_col in ["newly_added","closed","sold","licensed"]]
    for c in yes_no_cols:
        # 'Yes' => True, 'No' => False, anything else => None or leave as-is
        df[c] = df[c].apply(lambda x: True if str(x).strip().lower() == "yes" else
                                      False if str(x).strip().lower() == "no" else x)

    valid_db_cols = list(df.columns) 
    if not valid_db_cols:
        print(f"No matching columns found in '{excel_path}' for table '{table_name}'. Aborting.")
        return

    conn = connect_to_db()
    if not conn:
        print("Can't connect to DB. Aborting import.")
        return

    inserted_count = 0
    try:
        with conn.cursor() as cur:
            for row_idx, row_data in df.iterrows():
                placeholders = ",".join(["%s"] * len(valid_db_cols))
                col_str = ",".join(valid_db_cols)
                sql = f"INSERT INTO {table_name} ({col_str}) VALUES ({placeholders})"
                values = [row_data[col] if not pd.isna(row_data[col]) else None for col in valid_db_cols]
                cur.execute(sql, values)
                inserted_count += 1
        conn.commit()
        print(f"Inserted {inserted_count} rows into '{table_name}' from '{excel_path}'.")
    except Exception as e:
        print(f"Error inserting rows: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        
        if cmd == "import_excel" and len(sys.argv) >= 4:
            tbl = sys.argv[2]
            xls_path = sys.argv[3]
            import_excel_to_table(tbl, xls_path)
        
        else:
            print("Unknown or incomplete command. Examples:")
            print("  python dbconn.py import_excel <table_name> <excel_file>")
    else:
        print("** Database Utility Main **")
        print("Usage example:")
        print("  python dbconn.py import_excel ip_investmentsgroup MyFile.xlsx")
