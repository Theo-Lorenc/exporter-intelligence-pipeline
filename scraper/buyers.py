import requests
from bs4 import BeautifulSoup
from processing.text_utils import clean_text


HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


def collect_buyers():
    print("Searching for meat importers...")

    # ✅ You can expand this list later
    search_terms = [
        "beef importer singapore",
        "meat distributor singapore",
        "food import company singapore"
    ]

    buyers = []

    for term in search_terms:
        print(f"Searching: {term}")

        url = f"https://duckduckgo.com/html/?q={term.replace(' ', '+')}"
        resp = requests.get(url, headers=HEADERS)

        soup = BeautifulSoup(resp.text, "html.parser")

        for result in soup.select(".result__a"):
            name = clean_text(result.get_text())
            link = result.get("href")

            if link and name:
                buyers.append({
                    "buyer_name": name,
                    "buyer_website": link
                })

    return buyers