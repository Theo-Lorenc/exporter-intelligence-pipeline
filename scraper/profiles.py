from config.settings import *
from processing.text_utils import (
    clean_text,
    strip_boilerplate,
    filter_phone_candidates,
)
from scraper.website import extract_emails_from_website, safe_get

import json
import re
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed


# ✅ FIXED patterns
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_PATTERN = re.compile(r"\+?\d[\d\s\-\(\)]{7,}")


# ✅ helper
def unique_nonblank(values):
    return list(set([v.strip() for v in values if v and v.strip()]))


def extract_website_from_html(soup):
    link = soup.select_one("a[href*='http']")

    if not link:
        return ""

    href = link.get("href", "")

    if not href:
        return ""

    # ✅ clean HTML artifacts
    href = re.sub(r'["<>]', "", href)

    # ✅ remove trailing junk
    href = href.split(",")[0]

    if "mla.com.au" in href.lower():
        return ""

    return href.strip()


def is_valid_phone(phone):
    digits = re.sub(r"\D", "", phone)

    # ❌ remove known junk numbers
    if digits in ["39081678364"]:
        return False

    # ✅ realistic phone length
    return 8 <= len(digits) <= 12


# ✅ main parser
def parse_profile_page(html_text):
    soup = BeautifulSoup(html_text, "html.parser")

    page_text_raw = soup.get_text(" ", strip=True)
    page_text = strip_boilerplate(page_text_raw)

    text_lower = page_text.lower()

    # ✅ extract emails
    emails = EMAIL_PATTERN.findall(page_text)

    # ✅ remove MLA junk
    emails = [
        e for e in emails
        if "mla.com.au" not in e.lower()
    ]

    emails = unique_nonblank(emails)

    # ✅ extract phones first
    raw_phones = PHONE_PATTERN.findall(page_text)

    phones = filter_phone_candidates(raw_phones, page_text)

    # ✅ now apply validation filter
    phones = [p for p in phones if is_valid_phone(p)]

    # ✅ extract website
    website = extract_website_from_html(soup)

    # ✅ enrich emails via company website
    if website:
        emails.extend(extract_emails_from_website(website))
        emails = unique_nonblank(emails)

    # ✅ detect products (KEY FIX)
    product_families = []

    if "beef" in text_lower:
        product_families.append("Beef")

    if "lamb" in text_lower:
        product_families.append("Lamb")

    if "goat" in text_lower:
        product_families.append("Goat")

    if "wagyu" in text_lower:
        product_families.append("Wagyu Beef")

    # ✅ detect countries
    countries_served = []

    for country in ["singapore", "china", "japan", "usa", "uae"]:
        if country in text_lower:
            countries_served.append(country.title())

    return {
        "website": website,
        "emails": ", ".join(emails),
        "phones": ", ".join(phones),
        "product_families": ", ".join(product_families),
        "countries_served": ", ".join(countries_served),
        "description": page_text[:500]
    }


def clean_description(text, company_name):
    text = text.replace("☰", "")

    # Remove UI junk
    patterns = [
        "site insights",
        "admin view",
        "upload",
        "database table",
        "click on view report",
        "privacy consent",
    ]

    lowered = text.lower()

    for p in patterns:
        lowered = lowered.replace(p, " ")

    # Focus around company description
    if company_name.lower() in lowered:
        idx = lowered.find(company_name.lower())
        text = text[idx:idx+500]

    return text.strip()

def is_valid_phone(phone):
    digits = re.sub(r"\D", "", phone)

    # reject repeated junk numbers
    if digits in ["39081678364"]:
        return False

    return 8 <= len(digits) <= 12


# ✅ fetch
def fetch_profile_details(profile_url):
    if not profile_url:
        return {}

    response = safe_get(profile_url, timeout=REQUEST_TIMEOUT, retries=2)

    if not response:
        return {}

    return parse_profile_page(response.text)


# ✅ enrich row
def enrich_one_row(row):
    details = fetch_profile_details(row.get("profile_url"))

    # ✅ IMPORTANT: don't overwrite good data
    for key, value in details.items():
        if value:
            row[key] = value

    return row


# ✅ parallel processing
def enrich_rows(rows, max_workers=15):
    enriched = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(enrich_one_row, row): row
            for row in rows
        }

        for i, future in enumerate(as_completed(futures)):
            result = future.result()
            enriched.append(result)

            print(f"Profile {i+1}/{len(rows)}")

    return enriched