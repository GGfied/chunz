import os

import requests
from constants import FILE_DIR
from lxml import html
from parse_helpers import parse_clean_url, parse_append_hostname, parse_extract_datetime, parse_filename


def get_pages(category, year, long_month, page):
    url = 'https://www.mindef.gov.sg/web/wcm/connect/mindef/mindef-content/home'
    params = {
        'siteAreaName': 'mindef-content/home/news-and-events/latest-releases/{}/{}'.format(year, long_month),
        'srv': 'cmpnt',
        'selectedCategories': category,
        'cmpntid': 'dcb39e68-0637-4383-b587-29be9bb9bea5',
        'source': 'library',
        'cache': 'none',
        'contentcache': 'none',
        'connectorcache': 'none',
        'wcm_page.MENU-latest-releases': page,
    }
    page = requests.get(url, params=params)
    page_content = str(page.content)
    if page_content is None or len(page_content) < 10:
        return None

    tree = html.fromstring(page_content)

    links = tree.xpath('//a[@class="news-event-item-link"]/@href')
    links = list(map(parse_clean_url, links))
    links = list(map(parse_append_hostname, links))

    datetimes = tree.xpath('//span[@class="type-body-2"]/text()')
    datetimes = list(map(parse_extract_datetime, datetimes))
    datetimes = list(map(parse_filename, datetimes))

    num_links = len(links)
    num_dt = len(datetimes)
    i = 0
    pages = []

    while i < num_links and i < num_dt:
        pages.append({
            'link': links[i],
            'filename': datetimes[i],
            'category': category,
            'year': year,
            'month': long_month,
        })
        i += 1

    return pages


def get_month_pages(category, year, long_month):
    directory = os.path.join(FILE_DIR, category, str(year), long_month)

    if not os.path.exists(directory):
        os.makedirs(directory)

    month_pages = []

    for page_num in range(1, 100):
        pages = get_pages(category, year, long_month, page_num)
        if pages is None:
            break
        month_pages = month_pages + pages

    return month_pages


def get_year_pages(category, year):
    year_pages = []

    for month_str in [
        'january']:  # , 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']:
        month_pages = get_month_pages(category, year, month_str)
        year_pages = year_pages + month_pages

    year_pages.sort(key=lambda y: int(y['filename']))
    dup_map = dict()
    for i in range(len(year_pages)):
        filename = year_pages[i]['filename']
        dup_map[filename] = dup_map[filename] + 1 if filename in dup_map else 1
        dup_count = dup_map[filename]
        year_pages[i]['dup_prefix'] = '' if dup_count == 1 else '_{}'.format(dup_count)
        if dup_count == 2:
            year_pages[i - 1]['dup_prefix'] = '_1'

    with open(os.path.join(FILE_DIR, category, str(year), 'debug-listofpages.txt'), 'w') as f:
        f.write('\r\n'.join(list(map(str, year_pages))))

    return year_pages
