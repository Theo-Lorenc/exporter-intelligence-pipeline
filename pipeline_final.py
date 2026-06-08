# --- IMPORTS ---
import os
import sys
import re
import json
import time
import html
import random
import sqlite3
import subprocess
from datetime import datetime, UTC
from pathlib import Path
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# --- SESSION (performance + connection reuse) ---
SESSION = requests.Session()
SESSION.headers.update({"user-agent": "Mozilla/5.0"})
# Reuses TCP connections → faster scraping [1](https://stackoverflow.com/questions/76962976/understanding-the-session-in-pythons-requests-library)

# --- REGEX ---
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_PATTERN = re.compile(r"(?:\+?\d[\d\s\-()]{7,}\d)")

# --- HELPERS ---

def clean_text(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()


def unique_nonblank(items):
    seen, out = set(), []
    for x in items:
        x = clean_text(x)
        if x and x not in seen:
            out.append(x)
            seen.add(x)
    return out


# --- SAFE GET ---

def safe_get(url, retries=2, timeout=20):
    if not url:
        return None

    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    for _ in range(retries + 1):
        try:
            r = SESSION.get(url, timeout=timeout)
            if r.status_code == 200:
                return r
            if r.status_code in (403, 404):
                return None
        except Exception:
            pass

        time.sleep(random.uniform(1, 2))

    return None


# --- WEBSITE EXTRACTION (SMARTER) ---

def extract_website_from_details(details, page_text=""):
    # structured
    for k, v in details.items():
        if "website" in clean_text(k).lower():
            val = clean_text(v)
            val = val.rstrip(".,);")
            if not val.startswith(("http://", "https://")):
                val = "https://" + val
            return val

    # fallback regex
    match = re.search(r"(https?://\S+|www\.\S+)", page_text, re.IGNORECASE)
    if match:
        url = match.group(1).rstrip(".,);")
        if not url.startswith(("http://", "https://")):
            url = "https://" + url
        return url

    return ""


# --- WEBSITE EMAIL SCRAPING ---

def extract_emails_from_website(url):
    emails = []

    resp = safe_get(url)
    if not resp:
        return []

    html_text = resp.text
    emails.extend(EMAIL_PATTERN.findall(html_text))

    soup = BeautifulSoup(html_text, "html.parser")

    keywords = ["contact", "about", "team"]
    links = []

    for a in soup.select("a[href]"):
        href = a.get("href", "").lower()
        if any(k in href for k in keywords):
            links.append(urljoin(url, href))

    for link in unique_nonblank(links)[:2]:
        r = safe_get(link)
        if r:
            emails.extend(EMAIL_PATTERN.findall(r.text))

    return unique_nonblank(emails)


# --- PROFILE PARSER ---

def parse_profile_page(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    page_text = soup.get_text(" ", strip=True)

    emails = EMAIL_PATTERN.findall(page_text)
    phones = PHONE_PATTERN.findall(page_text)

    details = {}

    for row in soup.select("table tr"):
        cells = row.find_all(["th", "td"])
        if len(cells) == 2:
            k = clean_text(cells[0].text)
            v = clean_text(cells[1].text)
            if k and v:
                details[k] = v

    website = extract_website_from_details(details, page_text)

    for v in details.values():
        emails.extend(EMAIL_PATTERN.findall(v))

    emails.extend(extract_emails_from_website(website))

    return {
        "website": website,
        "emails": "; ".join(unique_nonblank(emails)),
        "phones": "; ".join(unique_nonblank(phones)),
        "details_json": json.dumps(details),
        "page_text_excerpt": page_text[:4000],
    }


# --- PROFILE FETCH (now uses safe_get) ---

def fetch_profile_details(url):
    r = safe_get(url)
    if not r:
        return {"profile_error": "failed fetch"}
    return parse_profile_page(r.text)


# --- DATABASE (FIXED WEBSITE + DEDUPE CONTACTS) ---

def build_database(df):
    conn = sqlite3.connect("exporters_final.db")
    cur = conn.cursor()

    cur.executescript("""
    DROP TABLE IF EXISTS companies;
    DROP TABLE IF EXISTS contacts;

    CREATE TABLE companies (
        id INTEGER PRIMARY KEY,
        name TEXT,
        profile_url TEXT UNIQUE,
        website TEXT,
        details_json TEXT
    );

    CREATE TABLE contacts (
        id INTEGER PRIMARY KEY,
        company_id INTEGER,
        email TEXT,
        phone TEXT,
        UNIQUE(company_id, email, phone)
    );
    """)

    for _, row in df.iterrows():
        cur.execute("""
        INSERT OR IGNORE INTO companies (name, profile_url, website, details_json)
        VALUES (?, ?, ?, ?)
        """, (
            row.get("company_name"),
            row.get("profile_url"),
            row.get("website"),
            row.get("details_json"),
        ))

        cur.execute("SELECT id FROM companies WHERE profile_url = ?", (row.get("profile_url"),))
        cid = cur.fetchone()[0]

        for e in row.get("emails", "").split(";"):
            cur.execute(
                "INSERT OR IGNORE INTO contacts (company_id, email, phone) VALUES (?, ?, '')",
                (cid, e)
            )

    conn.commit()
    conn.close()


# --- MAIN ---

def main():
    print("Running pipeline...")

    # Minimal stub (keeps your existing logic structure)
    print("Pipeline ready. Integrate rest of your listing + export logic as before.")

if __name__ == "__main__":
    main()