import os
import re

from docx.shared import Cm, Inches

FILE_DIR = os.path.dirname(os.path.abspath(__file__))
SINGLE_LINE_RE = re.compile(r'[\n\t\r]|[ ]{2,}|\xa0')
EXTRACT_DT_RE = re.compile(r'^.*?(\d{2}[ ][A-Z][a-z]{2}[ ]\d{4}).*$')
TEXT_ALIGN_IN_STYLE_RE = re.compile(r'text-align[ ]*:[ ]*([a-z]+)', flags=re.IGNORECASE)
HEIGHT_IN_STYLE_RE = re.compile(r'height[ ]*:[ ]*(\d+)', flags=re.IGNORECASE)
WIDTH_IN_STYLE_RE = re.compile(r'width[ ]*:[ ]*(\d+)', flags=re.IGNORECASE)
ARTICLE_TYPES_MAP = {
    'NEWS RELEASES': '001',
    'SPEECHES': '002',
    'OTHERS': '003',
}
TITLE_STYLE = 'TITLE_STYLE'
DATETIME_STYLE = 'DATETIME_STYLE'
CAPTION_STYLE = 'CAPTION_STYLE'
RUN_CAPTION_STYLE = 'RUN_CAPTION_STYLE'
BODY_STYLE = 'BODY_STYLE'
RUN_BODY_STYLE = 'RUN_BODY_STYLE'
MORE_RESOURCES_TITLE = 'More Resources'
MORE_RESOURCES_TITLE_STYLE = 'MORE_RESOURCES_TITLE_STYLE'
MORE_RESOURCES_LINK_STYLE = 'MORE_RESOURCES_LINK_STYLE'
LIST_BULLET_STYLE = 'LIST_BULLET_STYLE'
RUN_LIST_BULLET_STYLE = 'RUN_LIST_BULLET_STYLE'
LIST_NUMBER_STYLE = 'LIST_NUMBER_STYLE'
RUN_LIST_NUMBER_STYLE = 'RUN_LIST_NUMBER_STYLE'
TABLE_STYLE = 'TABLE_STYLE'
RUN_TABLE_STYLE = 'RUN_TABLE_STYLE'
RUN_LINK_STYLE = 'RUN_LINK_STYLE'

FONT_TNR = 'Times New Roman'
LOGO_URL = 'https://www.mindef.gov.sg/web/mindefstatic/themes/Portal8.5/images/logo_mindef.png'
DEFAULT_IMAGE_WIDTH = Inches(5.74)

COLSPAN = 'colspan'
ROWSPAN = 'rowspan'
HREF = 'href'
HEIGHT = 'height'
WIDTH = 'width'
STYLE = 'style'
BOLD_TAGS = ['b', 'strong']
ITALIC_TAGS = ['i', 'italic', 'em']
HANDLED_TAGS = ['p', 'div', 'span', 'hr', 'br', 'li', 'ol', 'ul', 'b', 'strong', 'em', 'italic', 'i', 'u', 'a',
                'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'sub', 'sup', 'section', 'article']
LIST_TAGS = ['ul', 'ol']
PARAGRAPH_TAGS = ['section', 'p', 'article']
REQ_NEW_PARA_TAGS = ['ul', 'ol', 'section', 'p', 'article']

LIST_TYPE_ORDERED = '5'
LIST_TYPE_UNORDERED = '1'

DEFAULT_CAPTION_PADDING = ''
