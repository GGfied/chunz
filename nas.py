!pip install lxml python-docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree, html
from datetime import datetime
import requests
import re
import unicodedata

SINGLE_LINE_RE = re.compile('[\\n\\t\\r]|[ ]{2,}')
EXTRACT_DT_RE = re.compile('^.+?(\d{2}[ ][A-Z][a-z]{2}[ ]\d{4}).+$')
ARTICLE_TYPES_MAP = {
    'NEWS RELEASES': '001',
    'SPEECHES': '002',
    'OTHERS': '003',
}
TITLE_STYLE = 'TITLE_STYLE'
DATETIME_STYLE = 'DATETIME_STYLE'
CAPTION_STYLE = 'CAPTION_STYLE'
BODY_STYLE = 'BODY_STYLE'
MORE_RESOURCES_TITLE_STYLE = 'MORE_RESOURCES_TITLE_STYLE'
MORE_RESOURCES_LINKS_STYLE = 'MORE_RESOURCES_LINKS_STYLE'
FONT_TNR = 'Times New Roman'
LOGO_URL = 'https://www.mindef.gov.sg/web/mindefstatic/themes/Portal8.5/images/logo_mindef.png'
VISITED_MAP = dict()
DEFAULT_PICTURE_WIDTH = Cm(15.24)

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
    
    more_resources_title_style = styles.add_style(MORE_RESOURCES_TITLE_STYLE, WD_STYLE_TYPE.PARAGRAPH)
    more_resources_title_style.font.name = FONT_TNR
    more_resources_title_style.font.size = Pt(10)
    more_resources_title_style.font.bold = True


def cleanup(txt):
#     return unicodedata.normalize("NFKD", re.sub(SINGLE_LINE_RE, '', txt)).strip()
    return re.sub(SINGLE_LINE_RE, '', txt).strip()


def extract_datetime(dt):
    return re.sub(EXTRACT_DT_RE, '\\1', dt)


def append_hostname(url):
    return 'https://www.mindef.gov.sg{}'.format(url) if url.startswith('/') else url


def parse_pr(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2013/january/2013Jan12-News-Releases-02121'):
    global VISITED_MAP

    if url in VISITED_MAP:
        return 'VISITED'
    else:
        VISITED_MAP[url] = True
        
    try:
        page = requests.get(url)
    except Exception as ex:
        print('-------------------------')
        print(url, ex)
        print('-------------------------')

        return
        
    tree = html.fromstring(page.content)
    
    titles = tree.xpath('//div[@class="article-detail__heading"]/div[contains(@class, "title")]/text()')
    title = cleanup(titles[0])
    
    article_types = tree.xpath('//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-label")]/text()')
    article_type = cleanup(article_types[0]).upper()
    
    datetime_strs = tree.xpath('//div[@class="article-detail__heading"]/div[@class="article-info"]/span[contains(@class, "item-published")]/text()')
    datetime_str = extract_datetime(datetime_strs[0])
    
    images = tree.xpath('//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//img/@src')
    for idx, img_url in enumerate(images):
        img_filename = fetch_image(img_url, idx)
        images[idx] = img_filename
        
    captions = tree.xpath('//article[contains(@class, "mindef-gallery-container")]/div[contains(@class, "mindef-gallery")]//span[@class="caption"]/text()')
    captions = list(map(cleanup, captions))
    
    body = tree.xpath('//div[@class="row"]//div[@class="article"]//p/text()')
    body = list(filter(lambda v: len(v) > 0, list(map(cleanup, body))))
#     for idx, b in enumerate(body):
#         body[idx] = etree.tostring(b, pretty_print=True)

    others_text = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/text()')
    others_text = list(map(cleanup, others_text))
    
    others_link = tree.xpath('//div[@class="more-resources"]/div[@class="more-resources__links"]/p/a/@href')
    others_link = list(map(append_hostname, others_link))

    print('Title', title)
    print('Article Type', article_type, ARTICLE_TYPES_MAP[article_type])
    print('DateTime', datetime_str)
    
    print('Images', images)
    
    print('Captions', captions)
    
    print('Body', body)
    
    print('Others Text', others_text)
    print('Others Link', others_link)
    
#     build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link)
    
    for ol in others_link:
        parse_pr(ol)

        
def fetch_image(url, idx):
    image_filename = 'img{}.png'.format(idx)
#     img_url = 'https://www.mindef.gov.sg{}'.format(url) if url.startswith('/') else url
#     image = requests.get(img_url)
#     open(image_filename, 'wb').write(image.content)

    return image_filename


def build_docx(article_type, title, datetime_str, images, captions, body, others_text, others_link):
    dup_prefix = ''
    filename = datetime.strptime(datetime_str, '%d %b %Y')
    print(filename)
    filename = 'MINDEF_{}{}.docx'.format(filename.strftime('%Y%m%d'), ARTICLE_TYPES_MAP[article_type], dup_prefix)
    print(filename)
    doc = Document()
    init_docx_styles(doc.styles)
    LOGO_FILENAME = fetch_image(LOGO_URL, 'LOGO')
    doc.add_picture(LOGO_FILENAME, width=DEFAULT_PICTURE_WIDTH)
    doc.add_paragraph(title, style=TITLE_STYLE)
    doc.add_paragraph(datetime_str, style=DATETIME_STYLE)
    num_images = len(images)
    num_captions = len(captions)
    num_overall = num_images if num_images > num_captions else num_captions
    if num_images > 0:
        doc.add_picture(images[0], width=DEFAULT_PICTURE_WIDTH)
    if num_captions > 0:
        doc.add_paragraph(captions[0], style=CAPTION_STYLE)
    for para in body:
        doc.add_paragraph(str(para), style=BODY_STYLE)
    for i in range(1, num_overall):
        doc.add_picture(images[i], width=DEFAULT_PICTURE_WIDTH)
        doc.add_paragraph(captions[i], style=CAPTION_STYLE)

    if len(others_text) > 0:
        doc.add_paragraph(MORE_RESOURCES_TITLE, style=MORE_RESOURCES_TITLE_STYLE)
        for i in others_text:
            doc.add_paragraph(MORE_RESOURCES_TITLE, style=MORE_RESOURCES_TITLE_STYLE)
            others_text[i]

    doc.save(filename)
    
    
parse_pr(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2019/june/13jun19_nr')
