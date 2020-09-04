import multiprocessing as mp
import os
import re

import requests
from lxml import html
from shared.constants import FILE_DIR, PARSE_PAGE_CATEGORY, PARSE_PAGE_YEAR, PARSE_PAGE_FILENAME, PARSE_PAGE_LINK, \
    PARSE_PAGE_MONTH, MONTHS, ERROR, DEFAULT_TIMEOUT_SECS
from shared.parse_helpers import parse_clean_url, parse_append_hostname, parse_extract_datetime, parse_filename, \
    parse_is_invalid_content, parse_get_datetimestr
from shared.writers import write_error


def get_page_model(link, filename, category, year, month):
    return {
        PARSE_PAGE_LINK: link,
        PARSE_PAGE_FILENAME: filename,
        PARSE_PAGE_CATEGORY: category,
        PARSE_PAGE_YEAR: year,
        PARSE_PAGE_MONTH: month,
    }


def get_page(link, directory='', nr_count_map=dict()):
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
    # get nr count subfix based on duplicates if 1st is 001, 2nd 002, 3rd 003, ...
    nr_count_map[filename] = nr_count_map[filename] + 1 if filename in nr_count_map else 1
    nr_count_subfix = nr_count_map[filename]
    filename = '{}00{}'.format(filename, nr_count_subfix)

    return get_page_model(link, filename, category='manual', year=year, month=long_month)


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
    page = requests.get(url, params=params, timeout=DEFAULT_TIMEOUT_SECS)

    if parse_is_invalid_content(page.content, page.status_code):
        return None

    tree = html.fromstring(page.content)

    links = tree.xpath('//a[@class="news-event-item-link"]/@href')
    links = list(map(parse_clean_url, links))
    links = list(map(parse_append_hostname, links))
    num_links = len(links)
    short_month_lc = long_month[0:3].lower()

    # invalid page
    if num_links > 0 and short_month_lc not in links[0].lower():
        return None

    datetimes = tree.xpath('//span[@class="type-body-2"]/text()')
    datetimes = list(map(parse_extract_datetime, datetimes))
    datetimes = list(map(parse_filename, datetimes))
    num_dt = len(datetimes)

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
    nr_count_map = dict()
    for i in range(len(year_pages)):
        filename = year_pages[i][PARSE_PAGE_FILENAME]
        # get nr count subfix based on duplicates if 1st is 001, 2nd 002, 3rd 003, ...
        nr_count_map[filename] = nr_count_map[filename] + 1 if filename in nr_count_map else 1
        nr_count_subfix = nr_count_map[filename]
        filename = '{}00{}'.format(filename, nr_count_subfix)
        year_pages[i][PARSE_PAGE_FILENAME] = filename

    with open(os.path.join(FILE_DIR, category, str(year), 'debug-listofpages.txt'), 'w') as f:
        f.write('\r\n'.join(list(map(str, year_pages))))

    return year_pages
