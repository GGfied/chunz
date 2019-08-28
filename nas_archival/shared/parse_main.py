import multiprocessing
import re

import requests
from lxml import html, etree
from shared.constants import ARTICLE_TYPES_MAP, EXT_DOCX, ERROR, MISSING_TYPE, NOT_SUPPORTED, DEFAULT_TIMEOUT_SECS
from shared.docx_helpers import docx_get_filename_prefix, docx_get_save_filename, docx_get_dup_prefix
from shared.docx_main import docx_build
from shared.parse_helpers import parse_append_hostname, parse_clean_url, parse_cleanup, parse_extract_img_link_caption, \
    parse_filename, parse_is_invalid_content, parse_get_datetimestr
from shared.writers import write_debug, write_details, write_error


def parse_article(url, filename='', dup_prefix='', directory='', visited_map=dict(), filename_to_dupcount_map=dict(),
                  is_follow_related_links=True, debug=False):
    if url in visited_map:
        return visited_map[url]

    print('Thread#:', multiprocessing.current_process())
    print('Processing: {}, {}'.format(filename, url))

    try:
        if not url.startswith('https://www.mindef.gov.sg/web/portal/mindef'):
            raise Exception('URL Not Supported')

        page = requests.get(url, timeout=DEFAULT_TIMEOUT_SECS)
        page_content = page.content

        if parse_is_invalid_content(page_content, page.status_code):
            raise Exception('Not Found, Status Code: {}'.format(page.status_code))
    except Exception as ex:
        print('-------------------------')
        print(url, ex)
        print('-------------------------')
        write_error(directory, error='Invalid URL: {}'.format(url), exception=ex)

        return ERROR

    tree = html.fromstring(page_content)

    if debug:
        titles = tree.xpath('//div[@class="article-detail__heading"]/div[contains(@class, "title")]/text()')
        title = parse_cleanup(titles[0])

        datetime_str = parse_get_datetimestr(tree)

        main = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]')

        if len(main) == 0:
            return

        main = main[0]
        children = main.getchildren()
        num_children = len(children)

        others_text = main.xpath('p/a')
        if len(others_text) == 0:
            others_text = main.xpath('ul/li/a')
            if len(others_text) > 0:
                print('missed out related text ul-li-a', title, datetime_str, url)
            elif num_children > 0:
                print('missed out unhandled related text', title, datetime_str, url, etree.tostring(main))
        others_text = list(map(lambda v: ''.join(v.itertext()), others_text))
        others_text = list(map(parse_cleanup, others_text))

        others_link = main.xpath('p/a/@href')
        if len(others_link) == 0:
            others_link = main.xpath('ul/li/a/@href')
            if len(others_link) > 0:
                print('missed out related link ul-li-a-@href', title, datetime_str, url)
            elif num_children > 0:
                print('missed out unhandled related link', title, datetime_str, url, etree.tostring(main))
        others_link = list(map(parse_clean_url, others_link))
        others_link = list(map(parse_append_hostname, others_link))

        visited_map[url] = True
        for i in range(len(others_link)):
            others_link[i] = parse_article(url=others_link[i], directory=directory, visited_map=visited_map,
                                           dup_prefix=dup_prefix, filename_to_dupcount_map=filename_to_dupcount_map,
                                           is_follow_related_links=is_follow_related_links, debug=debug)

        return

    titles = tree.xpath('//div[@class="article-detail__heading"]/div[contains(@class, "title")]/text()')
    title = parse_cleanup(titles[0])
    if not debug:
        write_details(directory, url, title)

    article_types = tree.xpath(
        '//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-label")]/text()')
    article_type = parse_cleanup(article_types[0] if len(article_types) > 0 else MISSING_TYPE).upper()

    if article_type not in ARTICLE_TYPES_MAP.keys():
        write_error(directory, error='Not Supported Article Type: {}, URL: {}'.format(article_type, url))

        return NOT_SUPPORTED

    datetime_str = parse_get_datetimestr(tree)

    images = tree.xpath(
        '//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//div[@class="item"]/figure')
    images = parse_extract_img_link_caption(images)

    body = tree.xpath('//div[@class="row"]//div[@class="article"]')[0]

    others_text = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a')
    others_text = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/ul/li/a') if len(others_text) == 0 else others_text
    others_text = list(map(lambda v: ''.join(v.itertext()), others_text))
    others_text = list(map(parse_cleanup, others_text))

    others_link = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/@href')
    others_link = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/ul/li/a/@href') if len(others_link) == 0 else others_link
    others_link = list(map(parse_clean_url, others_link))
    others_link = list(map(parse_append_hostname, others_link))

    article_type_prefix = ARTICLE_TYPES_MAP[article_type]

    write_debug(directory, msg='URL: {}'.format(url))
    write_debug(directory, msg='Title: {}'.format(title))
    write_debug(directory, msg='Article Type: {}, {}'.format(article_type, article_type_prefix))
    write_debug(directory, msg='DateTime: {}'.format(datetime_str))
    write_debug(directory, msg='Images: {}'.format(images))
    write_debug(directory, msg='Body: {}'.format(etree.tostring(body)))
    write_debug(directory, msg='Others Text: {}'.format(others_text))
    write_debug(directory, msg='Others Link: {}'.format(others_link))

    if not filename:
        filename = parse_filename(datetime_str)

    write_debug(directory, msg='Filename: {}'.format(filename))
    filename_prefix = docx_get_filename_prefix(filename, article_type_prefix, dup_prefix=dup_prefix)
    write_debug(directory, msg='Filename Prefix: {}'.format(filename))

    if filename_prefix in filename_to_dupcount_map:
        dup_count = filename_to_dupcount_map[filename_prefix]
        filename_to_dupcount_map[filename_prefix] += 1
        filename_prefix = '{}{}'.format(filename_prefix, docx_get_dup_prefix(dup_count))
    else:
        filename_to_dupcount_map[filename_prefix] = 1

    save_filename = docx_get_save_filename(filename_prefix, ext=EXT_DOCX)
    write_debug(directory, msg='Save Filename: {}'.format(save_filename))
    save_filename_no_ext = re.sub(EXT_DOCX + '$', '', save_filename)
    visited_map[url] = save_filename_no_ext

    if is_follow_related_links:
        for i in range(len(others_link)):
            others_link[i] = parse_article(url=others_link[i], directory=directory, visited_map=visited_map,
                                           dup_prefix=dup_prefix, filename_to_dupcount_map=filename_to_dupcount_map,
                                           is_follow_related_links=is_follow_related_links, debug=debug)

    docx_build(save_filename, filename_prefix, directory, title, datetime_str, images, body,
               others_text=others_text, others_link=others_link)

    return save_filename_no_ext
