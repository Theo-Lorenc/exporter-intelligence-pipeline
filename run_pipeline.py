import pandas as pd

from scraper.listings import collect_listings
from scraper.profiles import enrich_rows
from processing.text_utils import clean_text
from database.build_db import build_database
from output.excel_export import export_master_excel
from processing.brokerage_intelligence import assign_decision_category


def clean_dataframe(df):
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].fillna("").astype(str).map(clean_text)
    return df


# ✅ FUNCTION 1 (OUTSIDE main)
def compute_outreach_ready(row):
    has_email = bool(clean_text(row.get("emails")))
    has_website = bool(clean_text(row.get("website")))

    if has_email or has_website:
        return "Yes"
    return "No"


# ✅ FUNCTION 2 (OUTSIDE main)
def compute_contact_action(row):
    if row["outreach_ready"] == "Yes":
        return "Ready to contact"
    elif row["outreach_ready"] == "Maybe":
        return "Call or research email"
    else:
        return "Find contact details manually"


def main():
    print("Collecting listings...")
    listings = collect_listings()
    print(f"Collected {len(listings)} rows")

    print("Collecting profile details...")
    enriched = enrich_rows(listings)

    df = pd.DataFrame(enriched)
    df = clean_dataframe(df)

    def extract_product_flags(row):
        text = (
            clean_text(row.get("product_families", "")) + " " +
            clean_text(row.get("product_variants", ""))
        ).lower()

        return pd.Series({
            "has_beef": "Yes" if "beef" in text else "No",
            "has_lamb": "Yes" if "lamb" in text else "No",
            "has_goat": "Yes" if "goat" in text else "No"
        })

    # ✅ Apply it
    product_flags = df.apply(extract_product_flags, axis=1)
    df = pd.concat([df, product_flags], axis=1)

    def primary_product(row):
        if row["has_beef"] == "Yes":
            return "Beef"
        elif row["has_lamb"] == "Yes":
            return "Lamb"
        elif row["has_goat"] == "Yes":
            return "Goat"
        return "Unknown"

    df["primary_product"] = df.apply(primary_product, axis=1)


    # ✅ CORRECT ORDER
    df["outreach_ready"] = df.apply(compute_outreach_ready, axis=1)
    df["decision_category"] = df.apply(assign_decision_category, axis=1)
    df["contact_action"] = df.apply(compute_contact_action, axis=1)

    df["contact_status"] = "Not Contacted"
    df["last_contacted"] = ""
    df["response_received"] = "No"
    df["notes"] = ""

    rows = df.to_dict(orient="records")

    build_database(rows)
    export_master_excel()

    print("Pipeline complete ✅")


if __name__ == "__main__":
    main()
