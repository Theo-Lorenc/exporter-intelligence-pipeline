import pandas as pd
import sqlite3
from config.settings import DB_FILE, EXCEL_FILE


def export_master_excel():
    print("Exporting Excel...")

    conn = sqlite3.connect(DB_FILE)

    df = pd.read_sql_query("SELECT * FROM exporters", conn)

    conn.close()

    # ✅ Create Top Targets sheet
    df_top = df[
        (df["supplier_quality_score"] >= 7) &
        (df["outreach_ready"] == "Yes")
    ]

    # ✅ Write multiple sheets
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Exporters", index=False)
        df_top.to_excel(writer, sheet_name="Top Targets", index=False)

    print(f"Excel exported: {EXCEL_FILE}")