from config.settings import *
from processing.text_utils import clean_text, strip_boilerplate
from scraper.website import *

import json
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from processing.text_utils import clean_text, strip_boilerplate
from concurrent.futures import ThreadPoolExecutor, as_completed
import re

def parse_profile_page(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    page_text_raw = soup.get_text(" ", strip=True)
    page_text = strip_boilerplate(page_text_raw)

    def find_meta(names):
        for name in names:
            el = soup.select_one(f'meta[property="{name}"]') or soup.select_one(f'meta[name="{name}"]')
            if el and el.has_attr("content"):
                return strip_boilerplate(el["content"])
        return ""

    href_contacts = [
        a.get("href", "")
        for a in soup.select('a[href^="mailto:"], a[href^="tel:"]')
    ]
    href_emails = [
        x.replace("mailto:", "").strip()
        for x in href_contacts
        if x.startswith("mailto:")
    ]
    href_phones = [
        x.replace("tel:", "").strip()
        for x in href_contacts
        if x.startswith("tel:")
    ]

    regex_emails = EMAIL_PATTERN.findall(page_text)
    regex_phones = PHONE_PATTERN.findall(page_text)

    emails = unique_nonblank(href_emails + regex_emails)
    phones = filter_phone_candidates(
        href_phones + regex_phones,
        surrounding_text=page_text
    )

    details = {}

    for dt in soup.select("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            k = clean_text(dt.get_text(" ", strip=True))
            v = strip_boilerplate(dd.get_text(" ", strip=True))
            if k and v:
                details[k] = v

    for row in soup.select("table tr"):
        cells = row.find_all(["th", "td"])
        if len(cells) == 2:
            k = clean_text(cells[0].get_text(" ", strip=True))
            v = strip_boilerplate(cells[1].get_text(" ", strip=True))
            if k and v and k not in details:
                details[k] = v

    website = extract_website_from_details(details, page_text)
    countries_served = extract_countries_from_details(details)

    # extra emails from structured fields / raw HTML
    for v in details.values():
        emails.extend(EMAIL_PATTERN.findall(v))

    emails.extend(EMAIL_PATTERN.findall(html_text))

    # company website enrichment
    website_emails = extract_emails_from_website(website)
    emails.extend(website_emails)

    emails = unique_nonblank(emails)

    details_json = json.dumps(details, ensure_ascii=False) if details else ""

    return {
        "page_title": strip_boilerplate(soup.title.get_text(strip=True) if soup.title else ""),
        "page_heading": strip_boilerplate(
            soup.select_one("h1").get_text(" ", strip=True)
            if soup.select_one("h1") else ""
        ),
        "meta_description": find_meta(["description", "og:description"]),
        "meta_title": find_meta(["og:title"]),
        "website": website,
        "countries_served": countries_served,
        "emails": "; ".join(emails),
        "phones": "; ".join(phones),
        "details_json": details_json,
        "page_text_excerpt": page_text[:4000],
    }


def fetch_profile_details(profile_url):
    if not profile_url:
        return {}

    response = safe_get(profile_url, timeout=REQUEST_TIMEOUT, retries=2)
    if not response:
        return {"profile_error": f"Failed to fetch profile: {profile_url}"}

    return parse_profile_page(response.text)


def enrich_one_row(row):
    profile_url = row.get("profile_url", "")
    try:
        details = fetch_profile_details(profile_url) if profile_url else {}
    except Exception as e:
        details = {"profile_error": str(e)}

    merged = dict(row)
    merged.update(details)
    return merged


def enrich_rows(rows, max_workers=20):
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