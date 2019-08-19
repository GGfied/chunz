import re
import sys
from copy import deepcopy

import requests
from lxml import html

from shared.constants import URL_PARAMS, URL_PARAM_SITEAREANAME, HOME_URL, SITE_AREA_NAME, PARAM_PAGE, PARAM_CATEGORY, \
    CATEGORIES
from shared.parse_helpers import parse_append_hostname, parse_clean_url
from shared.writers import write_details, write_error

VISITED_MAP = dict()


def parse_links(links_objs=[], directory=''):
    for link_obj in links_objs:
        link = link_obj.attrib['href'] if 'href' in link_obj.attrib else ''
        linktext = link_obj.text

        if link in VISITED_MAP or link.startswith('javascript'):
            print('Skip Check Link: ', linktext, link)
            continue

        try:
            link = parse_clean_url(link)
            link = parse_append_hostname(link)
            print('Check Link: ', linktext, link)
            page = requests.get(link)
            status_code = int(page.status_code)

            if status_code != 200:
                print('Error: Status Code ({})'.format(status_code))
        except Exception as ex:
            print('-------------------------')
            print('Error Check Link: ', linktext, link, ex)
            print('-------------------------')
            write_error(directory, error='Invalid URL: {}'.format(link), exception=ex)

        VISITED_MAP[link] = True


def parse_article(link, directory=''):
    try:
        print('Check Article: ', link)
        page = requests.get(link)
        status_code = int(page.status_code)

        if status_code != 200:
            print('Error: Status Code ({})'.format(status_code))
    except Exception as ex:
        print('-------------------------')
        print(link, ex)
        print('-------------------------')
        write_error(directory, error='Invalid URL: {}'.format(link), exception=ex)

        return 'ERROR'

    tree = html.fromstring(page.content)
    links = tree.xpath('//div[@class="container article-detail"]//a')
    parse_links(links, directory)


def parse_list(siteareaname_l1, siteareaname_l2, param_page, param_category='', selected_category='', directory=''):
    url_params = deepcopy(URL_PARAMS)
    url_params[URL_PARAM_SITEAREANAME] = url_params[URL_PARAM_SITEAREANAME].format(l1=siteareaname_l1, l2=siteareaname_l2)

    if param_category is not '':
        url_params[param_category] = selected_category

    for page in range(1000000):
        url_params[param_page] = page

        print('Check List: ', url_params)
        page = requests.get(HOME_URL, params=url_params)
        page_content = str(page.content)
        if page_content is None or len(page_content) < 10:
            break

        tree = html.fromstring(page.content)
        links = tree.xpath('//a[@class="news-event-item-link"]/@href')

        for link in links:
            link = parse_clean_url(link)
            link = parse_append_hostname(link)
            parse_article(link, directory)
        sys.exit(1)


def parse_all(directory=''):
    for siteareaname_l1 in SITE_AREA_NAME.keys():
        for siteareaname_l2 in SITE_AREA_NAME[siteareaname_l1].keys():
            obj = SITE_AREA_NAME[siteareaname_l1][siteareaname_l2]
            for selected_category in obj[CATEGORIES]:
                parse_list(siteareaname_l1, siteareaname_l2, param_page=obj[PARAM_PAGE],
                           param_category=obj[PARAM_CATEGORY], selected_category=selected_category, directory=directory)

