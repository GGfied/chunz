# encoding: utf-8
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
    subprocess.call([sys.executable, "-m", "pip", "install", package])


if __name__ == '__main__':
    install('lxml')
    install('python-docx')
    install('requests')
    install('Pillow')

os.environ['OBJC_DISABLE_INITIALIZE_FORK_SAFETY'] = 'YES'

from docx import Document
from lxml import html

from shared import docxtopdf
from shared.constants import FILE_DIR, URL_PARAM_CATEGORY, GLOBAL_LOGO_FILENAME, GLOBAL_LOGO_PATH, PARSE_PAGE_CATEGORY, \
    PARSE_PAGE_YEAR, PARSE_PAGE_DUP_PREFIX, PARSE_PAGE_FILENAME, PARSE_PAGE_MONTH, PARSE_PAGE_LINK, \
    GLOBAL_SAVE_PDF_COUNTER, CPUS_TO_USE, PARSE_PAGE_SAVE_FOLDER

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
    <div class="article">\n            <p dir="ltr">A number of online initiatives have been planned for TD 2013 to reach out to Singaporeans. Centred on this year\'s campaign theme "Total Defence - Will You Stand With Me?", this year\'s online outreach comprises three key components: "Let\'s Stand Together" photo montage, "My Friend In Deed" contest as well as a 21-Day Challenge put out by bloggers. These online activities encourage and remind the public to stand with one another in good and bad times, be it overcoming challenges or doing something good for the community. The "My Friend In Deed" contest, in particular, prompts people to appreciate those who have stood with them to achieve remarkable feats or to overcome challenges.</p>\n\n<p dir="ltr"><strong>"Let\'s Stand Together" photo montage on campaign website</strong></p>\n\n<p dir="ltr">The key engagement activity on the official campaign website (<a href="http://www.standwithme.sg" target="_blank">www.standwithme.sg</a>) is the "Let\'s Stand Together" photo montage. The public will be invited to upload pictures of themselves, with their arms outstretched. These pictures will be linked to form the longest online human chain of Singaporeans "standing together". This simple activity is a visual representation of the campaign message that we can make a difference individually and collectively when we stand together as one.</p>\n\n<p dir="ltr"><strong>"My Friend in Deed" on Facebook</strong></p>\n\n<p dir="ltr">"My Friend In Deed" invites the public to submit stories (in the form of text, photos and/or videos) about how someone has stood with them to achieve a remarkable feat or to overcome a challenge. The contest will end on 22 March 2013, and participants will stand to win attractive prizes in the "Most Inspirational Story" category or "People\'s Choice" category. See <a href="http://www.facebook.com/ConnexionSG">www.facebook.com/ConnexionSG</a>.&#160;</p>\n\n<p dir="ltr"><strong>"The TD Challenge" App on Facebook</strong></p>\n\n<p dir="ltr">Members of the public will also get the opportunity to view the profiles of the teams participating in the TD Challenge and get live updates of the event on 15 February 2013 by logging on to <a href="http://www.facebook.com/ConnexionSG">www.facebook.com/ConnexionSG</a>. For more details about the TD Challenge, please refer to the fact sheet titled "TD Challenge".&#160;</p>\n\n<p dir="ltr"><strong>21-Day Challenge</strong></p>\n\n<p dir="ltr">Bloggers such as Dr Leslie Tay (<a href="http://ieatishootipost.sg" target="_blank">ieatishootipost.sg</a>) and Faz (<a href="http://thedramadiaries.com" target="_blank">thedramadiaries.com</a>) will blog, tweet or post on Facebook a challenge a day for their readers to complete. Through this, the readers will come to understand how the things that they do, individually and collectively, can help to foster a stronger sense of community and a greater sense of cohesion among Singaporeans. The reader who has completed a minimum of 10 challenges in the most creative way will stand to win attractive prizes. The 21-Day Challenge started on 26 January 2013 and will end on 15 February 2013.</p>\n\n<p dir="ltr">The list of 7 bloggers for this 21-Day Challenge are:</p>\n\n<p dir="ltr">&#160;</p>\n\n<table border="1" dir="ltr" style="width: 500px; table-layout: fixed; word-wrap: break-word; border-collapse : collapse; border-color : #696969;">\n\t<tbody>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; text-align: center; border-color : #696969;"><strong>Blogger</strong></td>\n\t\t\t<td style="overflow: hidden; width: 250px; text-align: center; border-color : #696969;"><strong>Blog Address</strong></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Dr Leslie Tay</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://ieatishootipost.sg/" title="null" target="_blank">http://ieatishootipost.sg/</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Faz</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://thedramadiaries.com/" target="_blank">http://thedramadiaries.com/</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Edmund Tay</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://edunloaded.com/">http://edunloaded.com/</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Diah</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://etrangle.net/">http://etrangle.net/</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Stephanie</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://stooffi.wordpress.com/">http://stooffi.wordpress.com/</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Jemimah</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://jemmawei.com" target="_blank">jemmawei.com</a></td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;">Que</td>\n\t\t\t<td style="overflow: hidden; width: 250px; border-color : #696969;"><a href="http://www.rynaque.com">www.rynaque.com</a></td>\n\t\t</tr>\n\t</tbody>\n</table>\n\n<p dir="ltr">&#160;</p>\n\n        </div>\
    </body>')
    xxxxx = html.fromstring(
        '<body><table><tr><td><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p><p style="text-align: center;"><em>阿興薄餅 Heng\'s Popiah</em></p></td></tr></table></body>')
    body = xxxx
    doc = Document()
    docx_init_styles(doc.styles)
    docx_build_body(body, doc, directory=os.path.join(FILE_DIR, 'debug'), filename_prefix='test')
    doc.save('debug/test.docx')
    docxtopdf.convert_to('debug', 'debug/test.docx')


def load_logo():
    GLOBALS[GLOBAL_LOGO_FILENAME] = GLOBAL_LOGO_PATH


def get_debug_dir():
    return os.path.join(FILE_DIR, 'debug')


def init_shared(args):
    GLOBALS[GLOBAL_SAVE_PDF_COUNTER] = args


def parse_page(page, is_follow_related_links=True, debug=False):
    if PARSE_PAGE_CATEGORY not in page and PARSE_PAGE_YEAR not in page and PARSE_PAGE_MONTH not in page:
        return

    directory = os.path.join(FILE_DIR, page[PARSE_PAGE_SAVE_FOLDER], page[PARSE_PAGE_CATEGORY], str(page[PARSE_PAGE_YEAR]), page[PARSE_PAGE_MONTH],
                             page[PARSE_PAGE_FILENAME] + page[PARSE_PAGE_DUP_PREFIX])

    if not debug and os.path.exists(directory):
        shutil.rmtree(directory)

    if not os.path.exists(directory):
        os.makedirs(directory)

    try:
        parse_article(url=page[PARSE_PAGE_LINK],
                      filename=page[PARSE_PAGE_FILENAME], dup_prefix=page[PARSE_PAGE_DUP_PREFIX],
                      directory=directory, is_follow_related_links=is_follow_related_links, debug=debug)
    except Exception:
        write_error(directory=directory, error='Exception', exception=traceback.format_exc())


def listbyyear(category, year, save_folder='', is_follow_related_links=True, debug=False):
    load_logo()
    year_pages = get_year_pages(category, year, save_folder=save_folder)
    mp_lock = mp.Value('i', 0)

    with mp.Pool(processes=CPUS_TO_USE, initializer=init_shared, initargs=(mp_lock,)) as p:
        res = [p.apply_async(parse_page, args=(page,),
                             kwds={'is_follow_related_links': is_follow_related_links, 'debug': debug}) for page in
               year_pages]
        p.close()
        p.join()
        res = [r.get() for r in res]


def parse_pages(urls=[], directory='', save_folder='manual', is_follow_related_links=True, debug=False):
    with mp.Pool(processes=CPUS_TO_USE) as p:
        pages = []
        res = p.map_async(partial(get_page, directory=directory, save_folder=save_folder, dup_map=dict()), urls, callback=pages.extend)
        res.wait()
        p.close()
        p.join()

    load_logo()
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
    parser.add_argument('--dont-follow-related-links', dest='follow-related-links', action='store_false',
                        help='(optional) DO NOT  Follow Related Links')
    parser.add_argument('--save-folder', dest='save-folder',
                        help='(optional) Save Folder for url(s) default is manual')
    parser.add_argument('--debug', dest='debug', action='store_true')
    args = vars(parser.parse_args())

    is_follow_related_links = args['follow-related-links']
    is_debug = args['debug']
    print(is_follow_related_links)

    if args['year'] is not None and args['category'] is not None:
        save_folder = args['save-folder'] or 'output_by_category'
        load_logo()
        listbyyear(category=args['category'], year=args['year'], save_folder=save_folder, debug=is_debug)
    elif args['url'] is not None:
        save_folder = args['save-folder'] or 'output_by_url'
        load_logo()
        debug_directory = get_debug_dir()
        init_shared(mp.Value('i', 0))
        parse_pages(urls=[args['url']], save_folder=save_folder, is_follow_related_links=is_follow_related_links, debug=is_debug)
    elif args['urls'] is not None:
        save_folder = args['save-folder'] or 'output_by_urls'
        load_logo()
        debug_directory = get_debug_dir()
        urls = args['urls'].split(',')
        init_shared(mp.Value('i', 0))
        parse_pages(urls=urls, save_folder=save_folder, is_follow_related_links=is_follow_related_links, debug=is_debug)
    elif is_debug:
        load_logo()
        debug_directory = get_debug_dir()
        docx_test()
    else:
        parser.print_help()


if __name__ == '__main__':
    start = time.time()
    docxtopdf.setup()
    main()
    end = time.time()
    print('Processing Time Taken: {}secs'.format(end - start))
