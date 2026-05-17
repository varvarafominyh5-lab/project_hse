!pip install requests beautifulsoup4 pandas openpyxl lxml
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
class RosstatParser:
    BASE_URL = "https://rosstat.gov.ru"

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

    def get_page(self, url: str) -> BeautifulSoup:
        response = self.session.get(url, timeout=10, verify=False)
        response.raise_for_status()
        return BeautifulSoup(response.text, "lxml")


    def find_excel_links(self, page_url: str) -> list[dict]:
        soup = self.get_page(page_url)
        links = []
        for tag in soup.find_all("a", href=True):
            href = tag["href"]
            if href.endswith((".xlsx", ".xls")):
                full_url = self.BASE_URL + href if href.startswith("/") else href
                links.append({
                    "name": tag.get_text(strip=True),
                    "url": full_url
                })
        return links

    def download_excel(self, url: str) -> pd.DataFrame:
        response = self.session.get(url, timeout=15, verify = False)
        response.raise_for_status()
        return pd.read_excel(io.BytesIO(response.content), header=None)

    def parse_labour_costs(self) -> list[dict]:
        url = f"{self.BASE_URL}/labour_costs"
        print(f"Ищем Excel-файлы на: {url}")
        links = self.find_excel_links(url)
        print(f"Найдено файлов: {len(links)}")
        for link in links:
            print(f"  - {link['name']}: {link['url']}")
        return links

