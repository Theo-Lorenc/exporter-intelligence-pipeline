import sqlite3
import pandas as pd
from config.settings import DB_FILE


def build_database(rows):
    print(f"Using database: {DB_FILE}")
    print("Building database...")

    df = pd.DataFrame(rows)

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

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
            notes TEXT
        )
    """)


    # ✅ Clear old data
    cur.execute("DELETE FROM exporters")

    # ✅ Insert new data
    for _, row in df.iterrows():
        cur.execute("""
        INSERT INTO exporters (
            company_name, description, website, emails, phones,
            countries_served, certifications, accreditations,
            product_families, product_variants,
            outreach_ready, decision_category, contact_action,
            contact_status, last_contacted, response_received, notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("company_name", ""),
        row.get("description", ""),
        row.get("website", ""),
        row.get("emails", ""),
        row.get("phones", ""),
        row.get("countries_served", ""),
        row.get("certifications", ""),
        row.get("accreditations", ""),
        row.get("product_families", ""),
        row.get("product_variants", ""),
        row.get("outreach_ready", ""),
        row.get("decision_category", ""),
        row.get("contact_action", ""),
        row.get("contact_status", ""),
        row.get("last_contacted", ""),
        row.get("response_received", ""),
        row.get("notes", "")
    ))

    conn.commit()
    conn.close()

    print("Database built successfully.")