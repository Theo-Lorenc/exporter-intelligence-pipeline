import re

BASE_URL = "https://www.aussiemeattradehub.com.au"
LIST_URL = BASE_URL + "/RMED/search/GetListingByPagination"
SEARCH_PAGE_URL = BASE_URL + "/rmed/search"

DB_FILE = "data/exporters.db"
EXCEL_FILE = "data/exporters.xlsx"
SCHEMA_SQL_FILE = "exporters_schema.sql"
SCHEMA_MD_FILE = "exporters_schema.md"
RUN_BAT_FILE = "run_pipeline.bat"

MAX_PAGES = 100
REQUEST_DELAY_SECONDS = 0.25
REQUEST_TIMEOUT = 60
MAX_WORKERS = 8

BOILERPLATE_PATTERNS = [
    r"ListingDetails",
    r"Exporters Database",
    r"Brand(?:,| &| &amp;) Licensing(?: & Assets| &amp; Assets)?",
    r"About the brand",
    r"Licensing program",
    r"Manage my licence",
    r"Assets",
    r"Global Insights",
    r"Trade Shows",
    r"Welcome,?",
    r"Report Centre",
    r"Login",
    r"Signup",
    r"Error Enquiry unable to send\.? Please try after sometime",
    r"Ok Processing This enquiry is being sent to the exporter - please wait",
    r"Successful Your enquiry was successfully(?: sent)?",
    r"×",
]

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_PATTERN = re.compile(r"(?:\+?\d[\d\s\-()]{7,}\d)")

HEADERS = {
    "accept": "application/json, text/javascript, */*; q=0.01",
    "content-type": "application/json; charset=UTF-8",
    "origin": BASE_URL,
    "referer": SEARCH_PAGE_URL,
    "x-requested-with": "XMLHttpRequest",
    "user-agent": "Mozilla/5.0",
}

SEARCH_PAYLOAD = {
    "search": {
        "SearchTerm": "",
        "SearchCategory": "All",
    }
}

FIELD_WEIGHTS = {
    "description": 1.00,
    "page_heading": 0.90,
    "meta_description": 0.70,
    "details_json": 0.60,
    "page_text_excerpt": 0.50,
}

KNOWN_CERTIFICATIONS = [
    "Certified Pasturefed",
    "Grainfed",
    "Halal",
    "MSA",
    "Organic",
    "Other 3rd Party Audited Verified",
]

KNOWN_ACCREDITATIONS = [
    "China",
    "EU",
    "Malaysia",
    "Egypt",
    "Indonesia",
    "Russia",
]

# Business-plan aligned settings
TARGET_COUNTRY = "Singapore"
COUNTRY_FIELD_KEYWORDS = [
    "countries",
    "country",
    "serviced",
    "services",
    "markets",
    "market",
    "countries serviced",
]
PREMIUM_PRODUCT_KEYWORDS = [
    "wagyu",
    "grainfed",
    "grain-fed",
    "msa",
    "premium beef",
]
SINGAPORE_FOCUS_PRODUCTS = [
    "beef",
    "grainfed beef",
    "wagyu beef",
]