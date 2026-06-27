import re
from config.settings import BOILERPLATE_PATTERNS


# ✅ basic text cleaner
def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


# ✅ remove generic boilerplate text
def strip_boilerplate(text):
    text = clean_text(text)

    for pattern in BOILERPLATE_PATTERNS:
        text = re.sub(pattern, " ", text, flags=re.IGNORECASE)

    return re.sub(r"\s+", " ", text).strip()


# ✅ remove duplicates + empty values safely
def unique_nonblank(values):
    seen = set()
    result = []

    for value in values:
        value = clean_text(value)

        if value and value not in seen:
            result.append(value)
            seen.add(value)

    return result


# ✅ clean phone numbers (important fix)
def filter_phone_candidates(candidates, surrounding_text=""):
    cleaned = []

    for raw in candidates:
        candidate = clean_text(raw)

        # ✅ remove ALL non-phone characters safely
        candidate = re.sub(r"[^\d+()\-\s]", "", candidate)

        digits = re.sub(r"\D", "", candidate)

        # ✅ must be realistic length
        if len(digits) < 8 or len(digits) > 15:
            continue

        # ✅ skip obvious junk (years, IDs)
        if digits.startswith("2021") or digits.startswith("2022"):
            continue

        cleaned.append(candidate)

    return unique_nonblank(cleaned)


# ✅ extract countries from structured fields
def extract_countries_from_details(details):
    chunks = []

    for k, v in details.items():
        key = clean_text(k).lower()

        if any(token in key for token in ["country", "market", "export"]):
            chunks.append(clean_text(str(v)))

    return " | ".join(unique_nonblank(chunks))