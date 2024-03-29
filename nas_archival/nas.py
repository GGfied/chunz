import argparse
import multiprocessing as mp
import os
import shutil
import subprocess
import sys
import time
import traceback
from functools import partial


def install(package):
    subprocess.call([sys.executable, "-m", "pip", "install", "--user", package])


if __name__ == '__main__':
    install('lxml')
    install('python-docx')
    install('requests')
    install('Pillow')
    install('docx2pdf')

os.environ['OBJC_DISABLE_INITIALIZE_FORK_SAFETY'] = 'YES'

from docx import Document
from lxml import html

# from shared import docxtopdf
from docx2pdf import convert
from shared.constants import FILE_DIR, URL_PARAM_CATEGORY, GLOBAL_LOGO_FILENAME, GLOBAL_LOGO_PATH, PARSE_PAGE_CATEGORY, \
    PARSE_PAGE_YEAR, PARSE_PAGE_FILENAME, PARSE_PAGE_MONTH, PARSE_PAGE_LINK, \
    GLOBAL_SAVE_PDF_COUNTER, CPUS_TO_USE, GLOBAL_SPEECH_LOGO_FILENAME, GLOBAL_SPEECH_LOGO_PATH

sys.path.append(FILE_DIR)

from shared.docx_body_table import docx_build_body
from shared.docx_main import docx_init_styles
from shared.globals import GLOBALS
from shared.parse_listing import get_year_pages, get_page
from shared.parse_main import parse_article
from shared.writers import write_error


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
            AFTER TEXT\
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
    <p>\
    How are you&nbsp;<b>dasdasdasd</b>&nbsp;sdfsdfsdfsd?\
    How are you&nbsp;<a href="http://wwww.google.com.sg">dasdasdasd</a>&nbsp;sdfsdfsdfsd?\
    </p>\
    <p>\
        <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
        <p>\
            <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
            <p>\
                <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
            </p>\
        </p>\
        <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
    </p>\
    <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
    <img alt="" src="/content/dam/imindef_media_library/photos/news_release/2016/jun/26jun16_nr/factsheet/thumbnail_26jun16_fs2.jpg" title="">\
    </body>')
    xxxxx = html.fromstring(
        '<body><table><tr><td><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p></td></tr></table></body>')
    body = xxxx
    doc = Document()
    docx_init_styles(doc.styles)
    docx_build_body(body, doc, directory=os.path.join(FILE_DIR, 'debug'), filename_prefix='test')
    doc.add_picture('debug/test.png')
    doc.save('debug/test.docx')
    convert(input_path='debug/test.docx', output_path='debug/test.pdf')
    # docxtopdf.convert_to('debug', 'debug/test.docx')


def load_logos():
    GLOBALS[GLOBAL_LOGO_FILENAME] = GLOBAL_LOGO_PATH
    GLOBALS[GLOBAL_SPEECH_LOGO_FILENAME] = GLOBAL_SPEECH_LOGO_PATH


def get_debug_dir():
    return os.path.join(FILE_DIR, 'debug')


def init_shared(args):
    GLOBALS[GLOBAL_SAVE_PDF_COUNTER] = args
    load_logos()


def parse_page(page, is_follow_related_links=True, debug=False):
    if PARSE_PAGE_CATEGORY not in page and PARSE_PAGE_YEAR not in page and PARSE_PAGE_MONTH not in page:
        return

    filename = page[PARSE_PAGE_FILENAME]
    directory = os.path.join(FILE_DIR, page[PARSE_PAGE_CATEGORY], str(page[PARSE_PAGE_YEAR]),
                             page[PARSE_PAGE_MONTH], filename)

    if os.path.exists(directory):
        shutil.rmtree(directory)

    os.makedirs(directory)

    try:
        parse_article(url=page[PARSE_PAGE_LINK], filename=filename, directory=directory,
                      is_follow_related_links=is_follow_related_links, debug=debug)
    except Exception:
        write_error(directory=directory, error='Exception', exception=traceback.format_exc())


def listbyyear(category, year, is_follow_related_links=True, debug=False, month_idx=None):
    load_logos()
    year_pages = get_year_pages(category, year, month_idx)
    mp_lock = mp.Value('i', 0)

    with mp.Pool(processes=CPUS_TO_USE, initializer=init_shared, initargs=(mp_lock,)) as p:
        res = [p.apply_async(parse_page, args=(page,),
                             kwds={'is_follow_related_links': is_follow_related_links, 'debug': debug}) for page in
               year_pages]
        p.close()
        p.join()
        res = [r.get() for r in res]


def parse_pages(urls=[], directory='', is_follow_related_links=True, debug=False):
    with mp.Pool(processes=CPUS_TO_USE) as p:
        pages = []
        res = p.map_async(partial(get_page, directory=directory, nr_count_map=dict()), urls, callback=pages.extend)
        res.wait()
        p.close()
        p.join()

    load_logos()
    mp_lock = mp.Value('i', 0)

    with mp.Pool(processes=CPUS_TO_USE, initializer=init_shared, initargs=(mp_lock,)) as p:
        res = [p.apply_async(parse_page, args=(page,),
                             kwds={'is_follow_related_links': is_follow_related_links, 'debug': debug}) for page in
               pages]
        p.close()
        p.join()
        res = [r.get() for r in res]


def main():
    parser = argparse.ArgumentParser(prog='NAS Archival', description='Parse URL to NAS .pdf')
    parser.add_argument('--year', dest='year', type=int,
                        help='year of <category> articles REQUIRED 1')
    parser.add_argument('--category', dest='category', choices=URL_PARAM_CATEGORY,
                        help='category of articles REQUIRED 1')
    parser.add_argument('--urls', dest='urls',
                        help='urls of articles, comma-separated REQUIRED 2')
    parser.add_argument('--url', dest='url',
                        help='url of article REQUIRED 3')
    parser.add_argument('--follow-related-links', dest='follow-related-links',
                        help='(optional) Follow Related Links?')
    parser.add_argument('--debug', dest='debug')
    parser.add_argument('--month', dest='month', type=int,
                        help='numeric month of <category> articles')
    args = vars(parser.parse_args())

    is_follow_related_links = not ('follow-related-links' in args and not args['follow-related-links'])

    if args['year'] is not None and args['category'] is not None:
        load_logos()
        month_idx = args['month'] - 1 if args['month'] else None
        listbyyear(category=args['category'], year=args['year'], month_idx=month_idx)
    elif args['url'] is not None:
        load_logos()
        debug_directory = get_debug_dir()
        init_shared(mp.Value('i', 0))
        parse_pages(urls=[args['url']], is_follow_related_links=is_follow_related_links, debug=False)
    elif args['urls'] is not None:
        load_logos()
        debug_directory = get_debug_dir()
        urls = args['urls'].split(',')
        init_shared(mp.Value('i', 0))
        parse_pages(urls=urls, is_follow_related_links=is_follow_related_links, debug=False)
    elif args['debug'] is not None:
        load_logos()
        debug_directory = get_debug_dir()
        docx_test()
    else:
        parser.print_help()


if __name__ == '__main__':
    start = time.time()
    # docxtopdf.setup()
    main()
    end = time.time()
    print('Processing Time Taken: {}secs'.format(end - start))
