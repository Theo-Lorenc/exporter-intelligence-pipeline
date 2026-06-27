import pandas as pd
from datetime import datetime

from scraper.listings import collect_listings
from scraper.profiles import enrich_rows
from scraper.buyers import collect_buyers

from processing.text_utils import clean_text
from processing.brokerage_intelligence import assign_decision_category

from database.build_db import build_database
from output.excel_export import export_master_excel


# ✅ clean dataframe
def clean_dataframe(df):
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].fillna("").astype(str).map(clean_text)
    return df


# ✅ outreach readiness (FIXED to avoid MLA emails)
def compute_outreach_ready(row):
    emails = clean_text(row.get("emails", ""))
    phones = clean_text(row.get("phones", ""))

    if emails and "mla.com.au" not in emails:
        return "Yes"

    if phones:
        return "Maybe"

    return "No"


def compute_contact_action(row):
    if row["outreach_ready"] == "Yes":
        return "Ready to contact"
    elif row["outreach_ready"] == "Maybe":
        return "Call or research email"
    else:
        return "Find contact details manually"


# ✅ improved scoring (FIXED)
def compute_supplier_quality(row):
    score = 0

    website = clean_text(row.get("website"))
    emails = clean_text(row.get("emails"))
    phones = clean_text(row.get("phones"))

    if website:
        score += 3

    if emails and "mla.com.au" not in emails:
        score += 4

    if phones:
        score += 1

    if clean_text(row.get("product_families")):
        score += 2

    return score


# ✅ targeting
def compute_target_priority(row):
    if row.get("has_lamb") == "Yes":
        if row["supplier_quality_score"] >= 7:
            return "High Priority Lamb Supplier"
        elif row["supplier_quality_score"] >= 4:
            return "Medium Priority Lamb Supplier"
        else:
            return "Low Priority Lamb Supplier"

    return "Not Target Product"


# ✅ product flags
def extract_product_flags(row):
    text = (
        clean_text(row.get("product_families", "")) + " " +
        clean_text(row.get("product_variants", ""))
    ).lower()

    return pd.Series({
        "has_beef": "Yes" if "beef" in text else "No",
        "has_lamb": "Yes" if "lamb" in text else "No",
        "has_goat": "Yes" if "goat" in text else "No",
    })


def compute_primary_product(row):
    if row["has_beef"] == "Yes":
        return "Beef"
    elif row["has_lamb"] == "Yes":
        return "Lamb"
    elif row["has_goat"] == "Yes":
        return "Goat"
    return "Unknown"


# ✅ matching logic (kept simple but correct)
def match_suppliers_to_buyers(suppliers_df, buyers_df):
    matches = []

    if buyers_df.empty:
        return pd.DataFrame(matches)

    for _, supplier in suppliers_df.iterrows():
        for _, buyer in buyers_df.iterrows():
            if supplier.get("has_beef") == "Yes":
                matches.append({
                    "supplier": supplier.get("company_name"),
                    "supplier_website": supplier.get("website"),
                    "buyer": buyer.get("buyer_name"),
                    "buyer_website": buyer.get("buyer_website"),
                })

    return pd.DataFrame(matches)


# ✅ MAIN PIPELINE
def main():
    print("Collecting listings...")
    listings = collect_listings()
    print(f"Collected {len(listings)} rows")

    print("Collecting profile details...")
    enriched = enrich_rows(listings)

    df = pd.DataFrame(enriched)
    df = clean_dataframe(df)

    # ✅ product detection FIRST
    product_flags = df.apply(extract_product_flags, axis=1)
    df = pd.concat([df, product_flags], axis=1)
    df["primary_product"] = df.apply(compute_primary_product, axis=1)

    # ✅ buyers
    print("Collecting buyers...")
    buyers = collect_buyers()
    buyers_df = pd.DataFrame(buyers)
    print(f"Found {len(buyers_df)} buyers")

    matches_df = match_suppliers_to_buyers(df, buyers_df)

    # ✅ business logic
    df["outreach_ready"] = df.apply(compute_outreach_ready, axis=1)
    df["supplier_quality_score"] = df.apply(compute_supplier_quality, axis=1)

    df["decision_category"] = df.apply(assign_decision_category, axis=1)
    df["contact_action"] = df.apply(compute_contact_action, axis=1)
    df["target_priority"] = df.apply(compute_target_priority, axis=1)

    # ✅ tracking columns
    df["contact_status"] = "Not Contacted"
    df["last_contacted"] = ""
    df["response_received"] = "No"
    df["notes"] = ""

    # ✅ persist + export
    rows = df.to_dict(orient="records")
    build_database(rows)
    export_master_excel(df, matches_df)

    print("Pipeline complete ✅")
    print(f"Last updated: {datetime.now()}")


if __name__ == "__main__":
    main()
