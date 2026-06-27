import requests
from bs4 import BeautifulSoup
from processing.text_utils import clean_text


HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


def collect_buyers():
    print("Loading known importers...")

    buyers = [
        {
            "buyer_name": "Bidfood Singapore",
            "buyer_website": "https://www.bidfood.com.sg"
        },
        {
            "buyer_name": "Angliss Singapore",
            "buyer_website": "https://www.angliss.com.sg"
        },
        {
            "buyer_name": "Indoguna Singapore",
            "buyer_website": "https://www.indoguna.com"
        },
        {
            "buyer_name": "Classic Fine Foods Singapore",
            "buyer_website": "https://www.classicfinefoods.com"
        },
        {
            "buyer_name": "FoodXervices Inc",
            "buyer_website": "https://www.foodxervices.com"
        }
    ]

    return buyers