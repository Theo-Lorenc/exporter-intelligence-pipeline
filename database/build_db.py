import sqlite3
import pandas as pd
from config.settings import DB_FILE


def build_database(rows):
    print(f"Using database: {DB_FILE}")
    print("Building database...")

    df = pd.DataFrame(rows)

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    # ✅ Create table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS exporters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT,
            description TEXT,
            website TEXT,
            emails TEXT,
            phones TEXT,
            countries_served TEXT,
            certifications TEXT,
            accreditations TEXT,
            product_families TEXT,
            product_variants TEXT,

            outreach_ready TEXT,
            decision_category TEXT,
            contact_action TEXT,
            contact_status TEXT,
            last_contacted TEXT,
            response_received TEXT,
            notes TEXT,
            supplier_quality_score INTEGER
        )
    """)

    # ✅ Clear old data
    cur.execute("DELETE FROM exporters")

    # ✅ Insert rows safely
    insert_query = """
        INSERT INTO exporters (
            company_name, description, website, emails, phones,
            countries_served, certifications, accreditations,
            product_families, product_variants,
            outreach_ready, decision_category, contact_action,
            contact_status, last_contacted, response_received, notes,
            supplier_quality_score
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    for _, row in df.iterrows():
        cur.execute(insert_query, (
            clean_value(row.get("company_name")),
            clean_value(row.get("description")),
            clean_value(row.get("website")),
            clean_value(row.get("emails")),
            clean_value(row.get("phones")),
            clean_value(row.get("countries_served")),
            clean_value(row.get("certifications")),
            clean_value(row.get("accreditations")),
            clean_value(row.get("product_families")),
            clean_value(row.get("product_variants")),
            clean_value(row.get("outreach_ready")),
            clean_value(row.get("decision_category")),
            clean_value(row.get("contact_action")),
            clean_value(row.get("contact_status")),
            clean_value(row.get("last_contacted")),
            clean_value(row.get("response_received")),
            clean_value(row.get("notes")),
            int(row.get("supplier_quality_score") or 0)
        ))

    conn.commit()
    conn.close()

    print("Database built successfully.")


# ✅ helper to clean DB values
def clean_value(value):
    if value is None:
        return ""
    return str(value).strip()