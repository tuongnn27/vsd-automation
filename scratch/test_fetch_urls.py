import requests
from bs4 import BeautifulSoup

session = requests.Session()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7',
    'Connection': 'keep-alive',
}

# Fetch a known "[OOPS!]" URL from the user's log
oops_url = "https://www.vsd.vn/vi/ad/193780"
print(f"Fetching OOPS URL: {oops_url}")
res = session.get(oops_url, headers=headers, timeout=15)
res.encoding = 'utf-8'
soup = BeautifulSoup(res.text, 'html.parser')
main = soup.find('main') or soup.find('article') or soup.find('div', class_='main-content') or soup.find('div', class_='content')
print(f"Status Code: {res.status_code}")
if main:
    print(f"Main text snippet: {main.get_text()[:300].strip()}")
else:
    print("No main div found")
    print(soup.get_text()[:300].strip())

print("\n------------------------------\n")

# Fetch a new URL that we just parsed (Item 0: 195963)
new_url = "https://www.vsd.vn/vi/ad/195963"
print(f"Fetching New URL: {new_url}")
res2 = session.get(new_url, headers=headers, timeout=15)
res2.encoding = 'utf-8'
soup2 = BeautifulSoup(res2.text, 'html.parser')
main2 = soup2.find('main') or soup2.find('article') or soup2.find('div', class_='main-content') or soup2.find('div', class_='content')
print(f"Status Code: {res2.status_code}")
if main2:
    print(f"Main text snippet: {main2.get_text()[:300].strip()}")
else:
    print("No main div found")
    print(soup2.get_text()[:300].strip())
