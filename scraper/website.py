import re
import requests
from urllib.parse import urljoin

from config.settings import BOILERPLATE_PATTERNS
from processing.text_utils import clean_text, unique_nonblank


# ✅ patterns
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
})

# ✅ caching
visited_websites = {}


# ✅ safe request handler
def safe_get(url, timeout=5, retries=2):
    if not url:
        return None

    url = clean_text(url).rstrip(".,);")

    if not url.startswith("http"):
        url = "https://" + url

    for _ in range(retries + 1):
        try:
            resp = SESSION.get(url, timeout=timeout)
            if resp.status_code == 200:
                return resp
            elif resp.status_code in [403, 404]:
                return None
        except Exception:
            pass

    return None


# ✅ normalise URLs
def normalize_website_url(url):
    url = clean_text(url)

    if not url:
        return ""

    if not url.startswith("http"):
        url = "https://" + url

    return url


# ✅ extract website from raw text
def extract_website_from_text(text):
    matches = re.findall(r"(https?://[^\s]+|www\.[^\s]+)", text or "", re.IGNORECASE)

    for match in matches:
        if "mla.com.au" not in match.lower():  # ✅ avoid directory links
            return normalize_website_url(match)

    return ""


# ✅ extract emails from company website
def extract_emails_from_website(url):
    if not url:
        return []

    if url in visited_websites:
        return visited_websites[url]

    emails = []

    response = safe_get(url)

    if not response:
        visited_websites[url] = emails
        return emails

    text = response.text

    raw_emails = EMAIL_PATTERN.findall(text)

    for email in raw_emails:
        cleaned = email.lower().strip()

        # ✅ skip MLA / directories / junk
        if any(bad in cleaned for bad in [
            "mla.com.au",
            "wixpress",
            "sentry",
            "example",
            "domain.com",
            ".jpg",
            ".png"
        ]):
            continue

        emails.append(cleaned)

    emails = unique_nonblank(emails)

    visited_websites[url] = emails
    return emails


# ✅ outreach readiness
def compute_outreach_ready(row):
    emails = clean_text(row.get("emails"))
    phones = clean_text(row.get("phones"))

    # ✅ only real emails count
    if emails and "@" in emails:
        return "Yes"

    # ✅ only reasonable phones
    if phones and len(re.sub(r"\D", "", phones)) >= 8:
        return "Maybe"

    return "No"