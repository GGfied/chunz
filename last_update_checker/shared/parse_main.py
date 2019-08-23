import re

import requests
from lxml import html
from shared.constants import ARTICLE_TYPES_MAP
from shared.docx_main import docx_build
from shared.parse_helpers import parse_append_hostname, parse_clean_url, parse_cleanup, parse_extract_img_link_caption, \
    parse_filename, parse_extract_datetime
from shared.writers import write_details, write_error


def parse_article(url, filename='', dup_prefix='', directory='', visited_map=dict(), dup_filename_map=dict(),
                  debug=False):
    if url in visited_map:
        return visited_map[url]

    print('Processing: {}, {}'.format(filename, url))

    try:
        page = requests.get(url)
    except Exception as ex:
        print('-------------------------')
        print(url, ex)
        print('-------------------------')
        write_error(directory, error='Invalid URL: {}'.format(url), exception=ex)

        return 'ERROR'

    tree = html.fromstring(page.content)

    titles = tree.xpath('//div[@class="article-detail__heading"]/div[contains(@class, "title")]/text()')
    title = parse_cleanup(titles[0])
    write_details(directory, url, title)

    article_types = tree.xpath(
        '//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-label")]/text()')
    article_type = parse_cleanup(article_types[0]).upper()

    datetime_strs = tree.xpath(
        '//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-published")]/text()')
    datetime_str = parse_extract_datetime(datetime_strs[0])

    images = tree.xpath(
        '//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//div[@class="item"]/figure')
    images = parse_extract_img_link_caption(images)

    body = tree.xpath('//div[@class="row"]//div[@class="article"]')[0]

    others_text = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/text()')
    others_text = list(map(parse_cleanup, others_text))

    others_link = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/@href')
    others_link = list(map(parse_clean_url, others_link))
    others_link = list(map(parse_append_hostname, others_link))

    # print('URL', url)
    # print('Title', title)
    # print('Article Type', article_type, ARTICLE_TYPES_MAP[article_type])
    # print('DateTime', datetime_str)
    # print('Images', images)
    # print('Body', body)
    # print('Others Text', others_text)
    # print('Others Link', others_link)

    if not filename:
        filename = parse_filename(datetime_str)
    # print('FILENAME', filename)
    filename_prefix = 'MINDEF_{}{}{}'.format(filename, ARTICLE_TYPES_MAP[article_type], dup_prefix)
    if filename_prefix in dup_filename_map:
        dup_filename_prefix = dup_filename_map[filename_prefix]
        dup_filename_map[filename_prefix] += 1
        filename_prefix = '{}_{}'.format(filename_prefix, dup_filename_prefix)
    else:
        dup_filename_map[filename_prefix] = 1

    save_filename = '{}.docx'.format(filename_prefix)
    # print('SAVE_FILENAME', save_filename)
    visited_map[url] = re.sub('.docx$', '.pdf', save_filename)

    if not debug:
        for i in range(len(others_link)):
            others_link[i] = parse_article(url=others_link[i], directory=directory, visited_map=visited_map,
                                           dup_prefix=dup_prefix, dup_filename_map=dup_filename_map)

    save_filename = docx_build(save_filename, filename_prefix, directory, title, datetime_str, images, body,
                               others_text, others_link)

    return re.sub('.docx$', '.pdf', save_filename)
