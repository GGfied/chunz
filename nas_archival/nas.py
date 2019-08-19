import os
import subprocess
import sys

from constants import LOGO_URL, FILE_DIR, TABLE_STYLE, RUN_TABLE_STYLE, RUN_BODY_STYLE, DEFAULT_IMAGE_WIDTH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Inches
from docx_helpers import docx_apply_text_align
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as WD_ALIGN_PARAGRAPH, WD_BREAK

sys.path.append(FILE_DIR)

from builddocx_body_table import docx_build_body, copy_run
from builddocx_main import docx_init_styles
from globals import GLOBALS
from listing import get_year_pages
from parse_helpers import parse_fetch_image
from parse_main import parse_article


def install(package):
    subprocess.call([sys.executable, "-m", "pip", "install", package])


if __name__ == '__main__':
    install('lxml')
    install('python-docx')
    install('requests')

from docx import Document
from lxml import html
import sys


def docx_test():
    xxx = html.fromstring('<html><head><title>sadasdas</title></head><body>\
        <div style="overflow:auto;"><table border="1" dir="ltr" style="width: 100%; border-collapse : collapse; border-color : #696969;">\
    <tbody>\
        <tr>\
            <td colspan="3" style="border-color: rgb(105, 105, 105); text-align: center;"><strong>Jury\'s Choice Awards</strong></td>\
        </tr>\
        <tr>\
            <td colspan="3" style="border-color: rgb(105, 105, 105); text-align: center;"><em><strong>The winners for the following awards were determined by a selection panel set up by the Organiser.</strong></em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); text-align: center; width: 223px;"><strong>Prizes</strong></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center; width: 328px;"><strong>Student Category</strong></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;"><strong>Open Category</strong></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 223px;">\
            <ul>\
                <li><strong>Feature Film Deal</strong></li>\
            </ul>\
            </td>\
            <td colspan="2" rowspan="1" style="border-color: rgb(105, 105, 105); width: 579px; text-align: center;">ciNE65 Movie Makers Award<br>\
            <br>\
            <em><strong>Winner:</strong> My Homeland: A Photography Project by Grandpa Chen</em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 223px;">\
            <ul>\
                <li><strong>Cash Prize $3,000</strong><br>\
                &nbsp;</li>\
                <li><strong>Learning Trip to an International Film Festival</strong><br>\
                &nbsp;</li>\
                <li><strong>Panasonic Professional 4K Camcorder (AG-UX180EN)&nbsp;</strong></li>\
            </ul>\
            </td>\
            <td style="border-color: rgb(105, 105, 105); width: 328px;">\
            <p style="text-align: center;">Overall Best Film Award</p>\
\
            <p style="text-align: center;"><em><strong>Winner: </strong>My Homeland: A Photography Project<br>\
            by Grandpa Chen</em></p>\
            </td>\
            <td style=" border-color : #696969;">\
            <p style="text-align: center;">Overall Best Film Award</p>\
\
            <p style="text-align: center;"><em><strong>Winner: </strong>$ingapura</em></p>\
            </td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 223px;">\
            <ul>\
                <li><strong>Cash Prize $1,000&nbsp;</strong><br>\
                &nbsp;</li>\
                <li><strong>Panasonic Professional 4K Camera (AG-UX90EN)</strong></li>\
            </ul>\
            </td>\
            <td style="border-color: rgb(105, 105, 105); width: 328px;">\
            <p style="text-align: center;">Best Cinematography</p>\
\
            <p style="text-align: center;"><em><strong>Winner:</strong> 村 Kampong</em></p>\
            </td>\
            <td style=" border-color : #696969;">\
            <p style="text-align: center;">Best Cinematography</p>\
\
            <p style="text-align: center;"><em><strong>Winner: </strong>一人一半</em></p>\
            </td>\
        </tr>\
        <tr>\
            <td colspan="1" rowspan="5" style="border-color: rgb(105, 105, 105); width: 223px;">\
            <ul>\
                <li><strong>Cash Prize $1,000&nbsp;</strong><br>\
                &nbsp;</li>\
                <li><strong>Panasonic product voucher worth $288 per award</strong></li>\
            </ul>\
            </td>\
            <td style="border-color: rgb(105, 105, 105); width: 328px; text-align: center;">Best Direction<br>\
            <br>\
            <em><strong>Winner: </strong>Broken</em></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;">Best Direction<br>\
            <br>\
            <em><strong>Winner: </strong>一人一半</em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 328px; text-align: center;">Best Editing<br>\
            <br>\
            <em><strong>Winner: </strong>Echoes of 1965</em></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;">Best Editing<br>\
            <br>\
            <em><strong>Winner: </strong>Ah Gong Garden</em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 328px; text-align: center;">Best Screenplay<br>\
            <br>\
            <em><strong>Winner:</strong> Built to Last</em></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;">Best Screenplay<br>\
            <br>\
            <em><strong>Winner: </strong>$ingapura</em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 328px; text-align: center;">Best Sound Design<br>\
            <br>\
            <em><strong>Winner: </strong>Pulau Ujong</em></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;">Best Sound Design<br>\
            <br>\
            <em><strong>Winner: </strong>Sound of Home</em></td>\
        </tr>\
        <tr>\
            <td style="border-color: rgb(105, 105, 105); width: 328px; text-align: center;">Best Art Direction<br>\
            <br>\
            <em><strong>Winner: </strong>chope</em></td>\
            <td style="border-color: rgb(105, 105, 105); text-align: center;">Best Art Direction<br>\
            <br>\
            <em><strong>Winner: </strong>一人一半</em></td>\
        </tr>\
    </tbody>\
</table></div>\
        <p><b><i><u>hypernest</u></i></b></p>\
        <p>ptext before<b>ptext bold</b><i>ptext italic</i>ptext after</p>\
        <p>ptext bold before<b>bold</b>ptext bold after</p>\
        <p><b>asdad</b> xxx <strong>strong</strong> xxx <b>sadas</b><br />xxx<br>xxx<i>italic</i>xxx<u>underline</u>xxx<ol><li>num 1</li><li>num 2</li></ol><ul><li>bullet 1</li><li>bullet 2</li></ul></p></body></html>')
    xxxx = html.fromstring('<body>\
    <table>\
    <tr><td style="text-align: center">Here is it \
    <ol><li>A</li><li>B</li></ol>\
    After <b>Here I am</b> \
    <div><p>asdasd</p></div>\
    <ul><li>A</li><li>B</li></ul></td></tr>\
    <tr><td><ul><li>A</li><li>B</li></ul></td></tr>\
    <tr><td style="width: 402px; border-color : #696969;">\
			<p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p>\
			<p style="text-align: center;"><em>Roots of Rejections</em></p>\
			<p style="text-align: center;"><em>From You to Me密切留忆</em></p>\
			<p style="text-align: center;"><em>A Portrait of Mum</em></p>\
			<p style="text-align: center;"><em>First Impression</em></p>\
			</td></tr>\
    <tr><td style="text-align: center; width: 417px; border-color : #696969;"><em><strong>Winner:</strong> Mum\'s Last Day at Work – MingEn Seafood</em></td></tr>\
    <tr><td style="width: 650px; border-color : #696969;">\
			<p><a href="https://www.youtube.com/watch?v=rLK8pZOP7RE" target="_blank" rel="noopener noreferrer">$ingapura</a></p>\
			<p>A slice of life piece that observes this society through a day in life of a driver. The driver, Ah Hock, is representative of a generation in Singapore who works mainly for survival, and to provide a better living for their family.&nbsp;&nbsp; &nbsp;</p>\
			</td></tr>\
    <tr><td style="width: 671px; border-color : #696969;">\
            <p>&nbsp;</p>\
            <p>sadasd</p>\
			<ul>\
				<li>Competition submission period&nbsp;</li>\
			</ul>\
			<p style="margin-left: 40px;">- Participants can submit their competition entries via the ciNE65 website (<a href="http://www.cine65.sg" target="_blank" rel="noopener noreferrer">www.cine65.sg</a>) from 9 November 2018 till 18&nbsp; March 2019, 1200hrs.<br>\
			<br>\
			- A total of 113 entries – 51 under the Open Category, and 62 under the Student Category – were received.</p>\
			</td></tr>\
    </table>\
    </body>')
    xxxxx = html.fromstring('<body><table><tr><td><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p></td></tr></table></body>')
    body = xxxxx.xpath('//body')[0]
    doc = Document()
    docx_init_styles(doc.styles)
    docx_build_body(body, doc)
    doc.save('debug/test.docx')


def load_logo():
    GLOBALS['LOGO_FILENAME'] = parse_fetch_image(url=LOGO_URL, idx='', filename_prefix='LOGO', directory=FILE_DIR)


def create_debug_dir():
    directory = os.path.join(FILE_DIR, 'debug')

    if not os.path.exists(directory):
        os.makedirs(directory)

    return directory


def listbyyear(category, year):
    load_logo()
    year_pages = get_year_pages(category, year)
    for page in year_pages:
        directory = os.path.join(FILE_DIR, page['category'], str(page['year']), page['month'],
                                 page['filename'] + page['dup_prefix'])

        if not os.path.exists(directory):
            os.makedirs(directory)
        parse_article(url=page['link'], filename=page['filename'], dup_prefix=page['dup_prefix'], directory=directory)


# load_logo()
# docx_test()
# listbyyear(category='news-releases', year=2013)
load_logo()
debug_directory = create_debug_dir()
# # parse_article(url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2013/september/2013Sep01-News-Releases-01938')
parse_article(
    url='https://www.mindef.gov.sg/web/portal/mindef/news-and-events/latest-releases/article-detail/2019/june/13jun19_fs',
    directory=debug_directory)
