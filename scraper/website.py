from config.settings import *
from processing.text_utils import clean_text
import re
import time
import random
import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from processing.text_utils import unique_nonblank
from config.settings import BARE_DOMAIN_PATTERN

SESSION = requests.Session()


SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
})


def safe_get(url, timeout=20, retries=2):
    if not url:
        return None

    url = clean_text(url).rstrip(".,);")
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    for _ in range(retries + 1):
        try:
            response = SESSION.get(url, timeout=timeout, allow_redirects=True)
            if response.status_code == 200:
                return response
            if response.status_code in {403, 404}:
                return None
        except Exception:
            pass
        time.sleep(random.uniform(1, 2))

    return None


def normalize_website_url(url):
    if not url:
        return ""

    url = url.strip()

    if not url.startswith("http"):
        url = "http://" + url

    return url


def extract_website_from_details(details, page_text=""):
    # 1) explicit website-like fields
    for k, v in details.items():
        key = clean_text(k).lower()
        val = clean_text(v)
        if val and ("website" in key or key == "web"):
            normalized = normalize_website_url(val)
            if normalized:
                return normalized

    # 2) explicit URLs in page text
    match = re.search(r"(https?://\S+|www\.\S+)", page_text, flags=re.IGNORECASE)
    if match:
        normalized = normalize_website_url(match.group(1))
        if normalized:
            return normalized

    # 3) bare-domain fallback
    match = BARE_DOMAIN_PATTERN.search(page_text)
    if match:
        candidate = match.group(0)
        if "@" not in candidate and not candidate.lower().endswith(
            (".jpg", ".jpeg", ".png", ".gif", ".pdf", ".svg", ".webp")
        ):
            normalized = normalize_website_url(candidate)
            if normalized:
                return normalized

    return ""


def extract_emails_from_website(url):
    if not url:
        return []

    visited = set()
    emails = set()

    # ✅ Pages to try (high success rate)
    keywords = [
        "",
        "contact",
        "contact-us",
        "about",
        "about-us",
        "team",
        "company",
        "sales",
        "export",
        "distributor"
    ]

    base = url.rstrip("/")

    for keyword in keywords:
        try:
            page_url = base if keyword == "" else f"{base}/{keyword}"

            if page_url in visited:
                continue

            visited.add(page_url)

            resp = safe_get(page_url)
            if not resp:
                continue

            text = resp.text

            # ✅ Extract emails
            found = EMAIL_PATTERN.findall(text)

            for email in found:
                cleaned = email.lower().strip()

                # ✅ Remove junk emails
                if any(bad in cleaned for bad in [
                    "example", "test", "placeholder",
                    "noreply", "no-reply"
                ]):
                    continue

                emails.add(cleaned)

        except Exception:
            continue

    # ✅ Also search for mailto links
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(text, "html.parser")

    for link in soup.find_all("a", href=True):
        href = link["href"]

        if "mailto:" in href:
            email = href.replace("mailto:", "").strip()
            emails.add(email.lower())    

    return list(emails)

def compute_outreach_ready(row):
    has_email = bool(clean_text(row.get("emails")))
    has_phone = bool(clean_text(row.get("phones")))

    if has_email:
        return "Yes"
    elif has_phone:
        return "Maybe"
    else:
        return "No"