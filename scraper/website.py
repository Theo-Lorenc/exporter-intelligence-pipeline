from config.settings import *
from processing.text_utils import clean_text
import re
import time
import random
import requests
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from processing.text_utils import unique_nonblank

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
    url = clean_text(url).rstrip(".,);")
    if not url:
        return ""
    if url.startswith("mailto:") or url.startswith("tel:"):
        return ""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        parsed = urlparse(url)
        if not parsed.netloc:
            return ""
    except Exception:
        return ""
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

    emails = []

    resp = safe_get(url)
    if not resp:
        return []

    html_text = resp.text
    emails.extend(EMAIL_PATTERN.findall(html_text))

    soup = BeautifulSoup(html_text, "html.parser")
    keywords = ["contact", "about", "team", "company"]
    links = []

    for a in soup.select("a[href]"):
        href = clean_text(a.get("href", ""))
        text = clean_text(a.get_text(" ", strip=True)).lower()
        if not href:
            continue
        href_lower = href.lower()
        if any(k in href_lower for k in keywords) or any(k in text for k in keywords):
            links.append(urljoin(url, href))

    for link in unique_nonblank(links)[:3]:
        resp2 = safe_get(link)
        if resp2:
            emails.extend(EMAIL_PATTERN.findall(resp2.text))

    return unique_nonblank(emails)