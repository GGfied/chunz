!pip install lxml python-docx
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree, html
from datetime import datetime
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
TITLE_STYLE = 'TITLE_STYLE'
DATETIME_STYLE = 'DATETIME_STYLE'
IMAGE_STYLE = 'IMAGE_STYLE'
CAPTION_STYLE = 'CAPTION_STYLE'
BODY_STYLE = 'BODY_STYLE'
MORE_RESOURCES_STYLE = 'MORE_RESOURCES_STYLE'
MORE_RESOURCES_LINKS_STYLE = 'MORE_RESOURCES_LINKS_STYLE'
FONT_TNR = 'Times New Roman'
LOGO_URL = 'https://www.mindef.gov.sg/web/mindefstatic/themes/Portal8.5/images/logo_mindef.png'
LOGO_FILENAME = fetch_image(LOGO_URL, 'LOGO')

def init_docx_styles(styles):
    title_style = styles.add_style(TITLE_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    title_style.font.name = FONT_TNR
    title_style.font.size = Pt(14)
    title_style.font.bold = True
    
    datetime_style = styles.add_style(DATETIME_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    datetime_style.font.name = FONT_TNR
    datetime_style.font.size = Pt(10)
    
    caption_style = styles.add_style(CAPTION_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    caption_style.font.name = FONT_TNR
    caption_style.font.size = Pt(10)
    
    body_style = styles.add_style(BODY_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    body_style.font.name = FONT_TNR
    body_style.font.size = Pt(12)

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
    
    for idx, img_url in enumerate(images):
        img_filename = fetch_image(img_url, idx)
        images[idx] = img_filename
    print('Images', images)
    
    print('Captions', captions)
    
    for idx, b in enumerate(body):
        body[idx] = etree.tostring(b, pretty_print=True)
    print('Body', body)
    
    print('Others Text', others_text)
    print('Others Link', others_link)
    
    build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link)
    
    for ol in others_link:
        parse_pr(ol)
        
def fetch_image(url, idx):
    image_filename = 'img{}.png'.format(idx)
    img_url = 'https://www.mindef.gov.sg{}'.format(url) if url.startswith('/') else url
    image = requests.get(img_url)
    open(image_filename, 'wb').write(image.content)

    return image_filename
        
def build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link):
    dup_prefix = ''
    filename = datetime.strptime(datetime_str, '%d %b %Y')
    print(filename)
    filename = 'MINDEF_{}{}.docx'.format(filename.strftime('%Y%m%d'), ARTICLE_TYPES_MAP[article_type], dup_prefix)
    print(filename)
    doc = Document()
    init_docx_styles(doc.styles)
    doc.add_picture(LOGO_FILENAME)
    doc.add_paragraph(title, style=TITLE_STYLE)
    doc.add_paragraph(datetime_str, style=DATETIME_STYLE)
    num_images = len(images)
    num_captions = len(captions)
    num_overall = num_images if num_images > num_captions else num_captions
    if num_images > 0:
        doc.add_picture(images[0])
    if num_captions > 0:
        doc.add_paragraph(captions[0], style=CAPTION_STYLE)
    for para in body:
        doc.add_paragraph(str(para), style=BODY_STYLE)
    for (var i = 1; i < num_overall; ++i):
        doc.add_picture(images[i])
        doc.add_picture(captions[i])

    doc.save(filename)
    
parse_pr(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2013/november/2013Nov19-News-Releases-02556')
