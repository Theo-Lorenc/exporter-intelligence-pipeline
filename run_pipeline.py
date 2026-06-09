from scraper.listings import collect_listings
from scraper.profiles import enrich_rows
from processing.text_utils import clean_text
from database.build_db import build_database
from output.excel_export import export_master_excel


def clean_dataframe(df):
    import pandas as pd
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].fillna("").astype(str).map(clean_text)
    return df


def main():
    print("Collecting listings...")
    listings = collect_listings()
    print(f"Collected {len(listings)} rows")

    print("Collecting profile details...")
    enriched = enrich_rows(listings)

    import pandas as pd
    df = pd.DataFrame(enriched)
    df = clean_dataframe(df)

    rows = df.to_dict(orient="records")

    build_database(rows)
    export_master_excel()

    print("Pipeline complete ✅")


if __name__ == "__main__":
    main()