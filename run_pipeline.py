import pandas as pd

from scraper import buyers
from scraper.listings import collect_listings
from scraper.profiles import enrich_rows
from processing.text_utils import clean_text
from database.build_db import build_database
from output.excel_export import export_master_excel
from processing.brokerage_intelligence import assign_decision_category
from scraper.buyers import collect_buyers


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

def compute_supplier_quality(row):
    score = 0

    # ✅ Has website
    if row.get("website"):
        score += 2

    # ✅ Has email (very important)
    if row.get("emails"):
        score += 4

    # ✅ Has phone
    if row.get("phones"):
        score += 1

    # ✅ Has product information
    if row.get("product_variants"):
        score += 2

    # ✅ Has certifications (strong trust signal)
    if row.get("certifications"):
        score += 3

    return score

def compute_target_priority(row):
    # Focus on lamb example
    if row.get("has_lamb") == "Yes":
        if row["supplier_quality_score"] >= 7:
            return "High Priority Lamb Supplier"
        elif row["supplier_quality_score"] >= 4:
            return "Medium Priority Lamb Supplier"
        else:
            return "Low Priority Lamb Supplier"

    return "Not Target Product"


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

    print("Collecting buyers...")
    buyers = collect_buyers()
    print(f"Found {len(buyers)} buyers")

    buyers_df = pd.DataFrame(buyers)

    def match_suppliers_to_buyers(suppliers_df, buyers_df):
        matches = []

        for _, supplier in suppliers_df.iterrows():
            for _, buyer in buyers_df.iterrows():
                if supplier.get("has_beef") == "Yes":
                    matches.append({
                        "supplier": supplier.get("company_name"),
                        "supplier_website": supplier.get("website"),
                        "buyer": buyer.get("buyer_name"),
                        "buyer_website": buyer.get("buyer_website")
                    })

        return pd.DataFrame(matches)

    matches_df = match_suppliers_to_buyers(df, buyers_df)

    # ✅ CORRECT ORDER
    df["outreach_ready"] = df.apply(compute_outreach_ready, axis=1)
    df["decision_category"] = df.apply(assign_decision_category, axis=1)
    df["contact_action"] = df.apply(compute_contact_action, axis=1)

    df["contact_status"] = "Not Contacted"
    df["last_contacted"] = ""
    df["response_received"] = "No"
    df["notes"] = ""
    df["supplier_quality_score"] = df.apply(compute_supplier_quality, axis=1)
    df["target_priority"] = df.apply(compute_target_priority, axis=1)

    rows = df.to_dict(orient="records")

    build_database(rows)
    export_master_excel(df, matches_df)

    print("Pipeline complete ✅")


if __name__ == "__main__":
    main()
