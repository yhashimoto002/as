import csv
import re
import sys
from urllib.parse import urljoin

from bs4 import BeautifulSoup
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

url_base = 'https://www.j-lis.go.jp/'
headers_ = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'}
retries_ = Retry(total=3,
                 backoff_factor=0.3,
                 status_forcelist=(500, 502, 503, 504))
timeout_ = 10
regex_prefecture = re.compile('都$|道央|道南|道北|道東|府$|県$')
regex_cities = re.compile('都$|道$|府$|県$|市$|区$|町$|村$')
result = {}


def get_link_from_html(link, regex):
    s = requests.Session()
    s.mount('https://', HTTPAdapter(max_retries=retries_))
    s.mount('http://', HTTPAdapter(max_retries=retries_))

    try:
        r = s.request('GET', urljoin(url_base, link), timeout=timeout_ , headers=headers_, )
    except OSError as e:
        print('OS Error: {}'.format(e))
        sys.exit()
    soup = BeautifulSoup(r.content, 'lxml')
    return soup.find_all('a', text=regex)


def get_response_code(link):
    print('Getting response code from {} ...'.format(link))
    s = requests.Session()
    s.mount('https://', HTTPAdapter(max_retries=retries_))
    s.mount('http://', HTTPAdapter(max_retries=retries_))

    try:
        r = s.request('GET', link, timeout=timeout_ , headers=headers_, )
    except:
        return 'Dead'
    return r.status_code


# scraping
print('Scraping link from site ...')
for link in get_link_from_html('spd/map-search/cms_1069.html', regex_prefecture):
    for link in get_link_from_html(link.get('href'), regex_cities):
        if link.string.strip() not in result:
            result[link.string.strip()] = [link.get('href'), get_response_code(link.get('href'))]


# output
print('Writing to csv ...')
try:
    with open('local_gov_urls.csv', 'w', encoding='utf-8_sig') as f:
        writer = csv.writer(f, lineterminator='\n')
        writer.writerow(['地方公共団体', 'URL', 'Status'])
        for key, value in result.items():
            city, (url, status) = key, value
            writer.writerow([city, url, status])
except OSError as e:
    print('OS Error: {}'.format(e))
else:
    print('Writing to csv is successfully ended.')
