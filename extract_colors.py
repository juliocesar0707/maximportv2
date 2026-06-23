import urllib.request
import re
from urllib.parse import urljoin

url = 'https://www.cloud.maxdata.com.br'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req).read().decode('utf-8', errors='ignore')
except Exception as e:
    print(f"Error fetching URL: {e}")
    exit(1)

css_links = re.findall(r'href=[\'"]([^\'"]+\.css)[\'"]', html)
all_css = ""

for link in css_links:
    css_url = urljoin(url, link)
    print(f"Fetching: {css_url}")
    try:
        req_css = urllib.request.Request(css_url, headers={'User-Agent': 'Mozilla/5.0'})
        css = urllib.request.urlopen(req_css).read().decode('utf-8', errors='ignore')
        all_css += css
    except Exception as e:
        print(f"Failed to fetch {css_url}: {e}")

colors = set(re.findall(r'#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}', html + all_css))
print("Colors found:")
for c in colors:
    print(c)
