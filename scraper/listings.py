from config.settings import *
import time
import html
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import requests
from processing.text_utils import strip_boilerplate, clean_text

SESSION = requests.Session()
SESSION.headers.update({
    "user-agent": HEADERS["user-agent"],
    "accept-language": "en-GB,en;q=0.9",
})

def fetch_list_page(page_number):
    payload = dict(SEARCH_PAYLOAD)
    payload["PageNumber"] = page_number

    response = SESSION.post(LIST_URL, headers=HEADERS, json=payload, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()

    try:
        data = response.json()
    except ValueError:
        data = response.text

    if isinstance(data, dict):
        html_snippet = data.get("Html") or data.get("html") or ""
    elif isinstance(data, str):
        html_snippet = data
    else:
        html_snippet = ""

    return html.unescape(html_snippet)


def parse_listing_html(html_snippet):
    soup = BeautifulSoup(html_snippet, "html.parser")
    rows = []

    for c in soup.select("div.company_list"):
        header_link = c.select_one("div.company_list_header a")
        desc_tag = c.select_one("div.company_list_body p")
        img_tag = c.select_one("img")

        rows.append({
            "company_name": header_link.get_text(strip=True) if header_link else "",
            "description": strip_boilerplate(desc_tag.get_text(" ", strip=True) if desc_tag else ""),
            "profile_url": urljoin(BASE_URL, header_link["href"]) if header_link and header_link.has_attr("href") else "",
            "image_url": urljoin(BASE_URL, img_tag["src"]) if img_tag and img_tag.has_attr("src") else "",
        })

    return rows


def collect_listings(max_pages=MAX_PAGES):
    all_rows = []

    for page in range(1, max_pages + 1):
        rows = parse_listing_html(fetch_list_page(page))
        print(f"Page {page}: {len(rows)}")

        if not rows:
            break

        all_rows.extend(rows)
        time.sleep(REQUEST_DELAY_SECONDS)

    return all_rows