import pandas as pd
import sqlite3
from config.settings import DB_FILE, EXCEL_FILE


def export_master_excel(df, matches_df):
    print("Exporting Excel...")

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:

        df.to_excel(writer, sheet_name="Suppliers", index=False)

        # ✅ Top targets
        df_top = df[
            (df["supplier_quality_score"] >= 7) &
            (df["outreach_ready"] == "Yes")
        ]
        df_top.to_excel(writer, sheet_name="Top Targets", index=False)

        # ✅ Matches sheet
        matches_df.to_excel(writer, sheet_name="Matches", index=False)

    print(f"Excel exported: {EXCEL_FILE}")
