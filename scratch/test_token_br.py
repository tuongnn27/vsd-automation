import requests
from bs4 import BeautifulSoup

session = requests.Session()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}

try:
    res = session.get("https://www.vsd.vn/vi/tin-thi-truong-co-so", headers=headers, timeout=10)
    print(f"Status: {res.status_code}")
    print(f"Content Length: {len(res.text)}")
    soup = BeautifulSoup(res.text, 'html.parser')
    meta = soup.find('meta', {'name': '__VPToken'})
    print(f"Meta token: {meta.get('content') if meta else None}")
except Exception as e:
    print(f"Error: {e}")
