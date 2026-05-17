!pip install requests beautifulsoup4 pandas openpyxl lxml
import requests
from bs4 import BeautifulSoup
import pandas as pd
import io
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
