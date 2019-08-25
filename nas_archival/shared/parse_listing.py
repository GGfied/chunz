import multiprocessing as mp
import os
import re
from datetime import datetime

import requests
from lxml import html
from shared.constants import FILE_DIR, PARSE_PAGE_CATEGORY, PARSE_PAGE_YEAR, PARSE_PAGE_FILENAME, PARSE_PAGE_LINK, \
    PARSE_PAGE_DUP_PREFIX, PARSE_PAGE_MONTH, MONTHS, ERROR, DEFAULT_TIMEOUT_SECS
from shared.docx_helpers import docx_get_dup_prefix
from shared.parse_helpers import parse_clean_url, parse_append_hostname, parse_extract_datetime, parse_filename, \
    parse_is_invalid_content, parse_get_datetimestr
from shared.writers import write_error


def get_page_model(link, filename, category, year, month, dup_prefix=''):
    return {
        PARSE_PAGE_LINK: link,
        PARSE_PAGE_FILENAME: filename,
        PARSE_PAGE_CATEGORY: category,
        PARSE_PAGE_YEAR: year,
        PARSE_PAGE_MONTH: month,
        PARSE_PAGE_DUP_PREFIX: dup_prefix,
    }


def get_page(link, directory='', folder='manual', dup_map=dict()):
    print('Info Processing: {}'.format(link))
    res = re.search(r'/(\d{4})/([^/]+?)/', link)

    if not res:
        write_error(directory, error='Invalid URL No Year or Month: {}'.format(link))

        return ERROR

    year = int(res.group(1))
    long_month = res.group(2)

    try:
        page = requests.get(link, timeout=DEFAULT_TIMEOUT_SECS)

        if parse_is_invalid_content(page.content, page.status_code):
            raise Exception('Content: {}, Error Code: {}'.format(page.content, page.status_code))
    except Exception as ex:
        write_error(directory, error='Invalid URL: {}'.format(link), exception=ex)

        return ERROR

    tree = html.fromstring(page.content)

    datetime_str = parse_get_datetimestr(tree)
    filename = parse_filename(datetime_str)
    dup_prefix = ''

    if filename not in dup_map:
        dup_map[filename] = 0
    else:
        dup_map[filename] = dup_map[filename] + 1
        dup_prefix = docx_get_dup_prefix(dup_map[filename])

    return get_page_model(link, filename, category=folder, year=year, month=long_month, dup_prefix=dup_prefix)


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
    page_html = requests.get(url, params=params, timeout=DEFAULT_TIMEOUT_SECS)

    if parse_is_invalid_content(page_html.content, page_html.status_code):
        return None

    tree = html.fromstring(page_html.content)

    links = tree.xpath('//a[@class="news-event-item-link"]/@href')
    links = list(map(parse_clean_url, links))
    links = list(map(parse_append_hostname, links))

    datetimes = tree.xpath('//span[@class="type-body-2"]/text()')
    datetimes = list(map(parse_extract_datetime, datetimes))
    datetimes = list(map(parse_filename, datetimes))

    num_links = len(links)
    num_dt = len(datetimes)
    month_num = datetime.strptime(long_month, '%B').strftime('%m')
    yaer_month_prefix = '{}{}'.format(year, month_num)

    if page == 1 and num_dt > 0 and not str(datetimes[0]).startswith(str(yaer_month_prefix)):
        return None

    i = 0
    pages = []

    while i < num_links and i < num_dt:
        pages.append(get_page_model(link=links[i], filename=datetimes[i], category=category,
                                    year=year, month=long_month))
        i += 1

    return pages


def get_month_pages(category, year, long_month):
    print('Retrieving Articles from', long_month, year)
    directory = os.path.join(FILE_DIR, category, str(year), long_month)

    if not os.path.exists(directory):
        os.makedirs(directory)

    month_pages = []

    for page_num in range(1, 100000):
        pages = get_pages(category, year, long_month, page_num)
        if pages is None:
            break
        month_pages = month_pages + pages

    print('Retrieved {} Articles from'.format(len(month_pages)), long_month, year)
    return month_pages


def get_year_pages(category, year):
    year_pages = []
    with mp.Pool(processes=4) as p:
        res = [p.apply_async(get_month_pages, args=(category, year, month_str)) for month_str in MONTHS]
        p.close()
        p.join()
        for month_pages_res in res:
            year_pages = year_pages + month_pages_res.get()

    year_pages.sort(key=lambda y: int(y[PARSE_PAGE_FILENAME]))
    dup_map = dict()
    for i in range(len(year_pages)):
        filename = year_pages[i][PARSE_PAGE_FILENAME]
        dup_map[filename] = dup_map[filename] + 1 if filename in dup_map else 1
        dup_count = dup_map[filename]
        year_pages[i][PARSE_PAGE_DUP_PREFIX] = '' if dup_count == 1 else docx_get_dup_prefix(dup_count)
        if dup_count == 2:
            year_pages[i - 1][PARSE_PAGE_DUP_PREFIX] = docx_get_dup_prefix(1)

    with open(os.path.join(FILE_DIR, category, str(year), 'debug-listofpages.txt'), 'w') as f:
        f.write('\r\n'.join(list(map(str, year_pages))))

    return year_pages
