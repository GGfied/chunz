!pip install lxml python-docx
from docx import Document
from lxml import etree, html
from datetime import datetime
from IPython.display import HTML
import requests
import json
import re

SINGLE_LINE_RE = re.compile('[\\n\\t\\r]|[ ]{2,}')
EXTRACT_DT_RE = re.compile('^.+?(\d{2}[ ][A-Z][a-z]{2}[ ]\d{4}).+$')
ARTICLE_TYPES_MAP = {
    'NEWS RELEASES': '001',
    'SPEECHES': '002',
    'OTHERS': '003',
}

def cleanup(txt):
    return re.sub(SINGLE_LINE_RE, '', txt)

def extract_datetime(dt):
    return re.sub(EXTRACT_DT_RE, '\\1', dt)

def parse_pr(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2013/january/2013Jan12-News-Releases-02121'):
    try:
        page = requests.get(url)
    except Exception as ex:
        print('-------------------------')
        print(url, ex)
        print('-------------------------')

        return
        
    tree = html.fromstring(page.content)
    title = cleanup(tree.xpath('//div[@class="article-detail__heading"]/div[contains(@class, "title")]/text()')[0])
    article_type = cleanup(tree.xpath('//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-label")]/text()')[0]).upper()
    datetime_str = extract_datetime(tree.xpath('//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-published")]/text()')[0])
    images = tree.xpath('//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//img/@src')
    captions = tree.xpath('//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//span[@class="caption"]/text()')
    body = tree.xpath('//div[@class="row"]//div[@class="article"]//p')
    others_text = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/text()')
    others_link = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/@href')

    print('Title', title)
    print('Article Type', article_type, ARTICLE_TYPES_MAP[article_type])
    print('DateTime', datetime_str)
    print('Images', images)
    print('Captions', captions)
    for b in body:
        print('Body', etree.tostring(b, pretty_print=True))
    print('Others Text', others_text)
    print('Others Link', others_link)
    
    build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link)
    
    for ol in others_link:
        parse_pr(ol)
        
def build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link):
    dup_prefix = ''
    filename = datetime.strptime(datetime_str, '%d %b %Y')
    print(filename)
    filename = 'MINDEF_{}{}.docx'.format(filename.strftime('%Y%m%d'), ARTICLE_TYPES_MAP[article_type], dup_prefix)
    print(filename)
    doc = Document()
    doc.add_heading(title, 0)
    doc.save(filename)
    
    
    
parse_pr(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2013/november/2013Nov19-News-Releases-02556')
