import requests
from bs4 import BeautifulSoup
import json
import re

session = requests.Session()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'vi-VN,vi;q=0.9',
    'Connection': 'keep-alive',
}

# First get the page to extract VPToken
res = session.get("https://www.vsd.vn/vi/tin-thi-truong-co-so", headers=headers, timeout=15)
soup = BeautifulSoup(res.text, 'html.parser')
meta = soup.find('meta', {'name': '__VPToken'})
vptoken = meta.get('content') if meta else None
print(f"VPToken: {vptoken}")

if vptoken:
    # Use AJAX POST to get page 1
    ajax_headers = {
        'User-Agent': 'Mozilla/5.0',
        'Content-Type': 'application/json;charset=utf-8',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': 'https://www.vsd.vn/vi/tin-thi-truong-co-so',
        'Origin': 'https://www.vsd.vn',
        '__VPToken': vptoken
    }
    payload = {'SearchKey': 'TCPH', 'CurrentPage': 1}
    response = session.post("https://www.vsd.vn/vi/tin-thi-truong-co-so", headers=ajax_headers, json=payload, timeout=15)
    
    ajax_soup = BeautifulSoup(response.text, 'html.parser')
    news_items = ajax_soup.find_all('li')
    print(f"Total items on page 1: {len(news_items)}")
    for i, item in enumerate(news_items[:10]):
        h3 = item.find('h3')
        if h3:
            link = h3.find('a')
            if link:
                print(f"Item {i}: Title: {link.get_text(strip=True)}, Href: {link.get('href')}")
