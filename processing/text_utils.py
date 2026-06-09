import re
from config.settings import *

def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def strip_boilerplate(text):
    text = clean_text(text)
    for pattern in BOILERPLATE_PATTERNS:
        text = re.sub(pattern, " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def unique_nonblank(values):
    seen = set()
    out = []
    for value in values:
        value = clean_text(value)
        if value and value not in seen:
            out.append(value)
            seen.add(value)
    return out