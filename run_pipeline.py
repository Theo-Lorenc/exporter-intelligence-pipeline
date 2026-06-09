from scraper.listings import collect_listings
from scraper.profiles import enrich_rows
from processing.text_utils import clean_text
from database.build_db import build_database
from output.excel_export import export_master_excel
from processing.brokerage_intelligence import assign_decision_category


def clean_dataframe(df):
    import pandas as pd
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].fillna("").astype(str).map(clean_text)
    return df


def compute_outreach_ready(row):
    has_email = bool(clean_text(row.get("emails")))
    has_website = bool(clean_text(row.get("website")))

    if has_email or has_website:
        return "Yes"
    return "No"


def main():
    print("Collecting listings...")
    listings = collect_listings()
    print(f"Collected {len(listings)} rows")

    print("Collecting profile details...")
    enriched = enrich_rows(listings)

    df = pd.DataFrame(enriched)
    df = clean_dataframe(df)

    # ✅ Apply outreach readiness
    df["outreach_ready"] = df.apply(compute_outreach_ready, axis=1)

    # ✅ Apply decision categories
    df["decision_category"] = df.apply(assign_decision_category, axis=1)

    # ✅ Contact tracking fields
    df["contact_status"] = "Not Contacted"
    df["last_contacted"] = ""
    df["response_received"] = "No"
    df["notes"] = ""

    # ✅ -------------------------------
    # Add outreach readiness logic
    # ✅ -------------------------------
    def compute_outreach_ready(row):
        has_email = bool(clean_text(row.get("emails")))
        has_website = bool(clean_text(row.get("website")))

        if has_email or has_website:
            return "Yes"
        return "No"

    df["outreach_ready"] = df.apply(compute_outreach_ready, axis=1)

    # ✅ -------------------------------
    # Add decision category
    # ✅ -------------------------------
    df["decision_category"] = df.apply(assign_decision_category, axis=1)

    # ✅ -------------------------------
    # Add contact tracking (CRM fields)
    # ✅ -------------------------------
    df["contact_status"] = "Not Contacted"
    df["last_contacted"] = ""
    df["response_received"] = "No"
    df["notes"] = ""

    # ✅ Convert to rows
    rows = df.to_dict(orient="records")

    build_database(rows)
    export_master_excel()

    print("Pipeline complete ✅")


if __name__ == "__main__":
    main()