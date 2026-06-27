import pandas as pd
from config.settings import EXCEL_FILE


def export_master_excel(df, matches_df):
    print("Exporting Excel...")

    # ✅ Ensure missing columns don’t crash filtering
    if "supplier_quality_score" not in df.columns:
        df["supplier_quality_score"] = 0

    if "outreach_ready" not in df.columns:
        df["outreach_ready"] = ""

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:

        # ✅ Main suppliers sheet
        df.to_excel(writer, sheet_name="Suppliers", index=False)

        # ✅ Top targets (safe filtering)
        df_top = df[
            (df["supplier_quality_score"] >= 7) &
            (df["outreach_ready"] == "Yes")
        ].sort_values(by="supplier_quality_score", ascending=False)

        df_top.to_excel(writer, sheet_name="Top Targets", index=False)

        # ✅ Matches sheet (safe fallback)
        if matches_df is not None and not matches_df.empty:
            matches_df.to_excel(writer, sheet_name="Matches", index=False)

    print(f"Excel exported: {EXCEL_FILE}")