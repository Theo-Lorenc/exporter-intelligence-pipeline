import pandas as pd
import sqlite3
from config.settings import DB_FILE, EXCEL_FILE


def export_master_excel():
    print("Exporting Excel...")

    conn = sqlite3.connect(DB_FILE)

    df = pd.read_sql_query("SELECT * FROM exporters", conn)

    conn.close()

    # ✅ Basic clean export
    df.to_excel(EXCEL_FILE, index=False)

    print(f"Excel exported: {EXCEL_FILE}")