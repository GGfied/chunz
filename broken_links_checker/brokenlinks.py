import os
import subprocess
import sys

from shared.constants import FILE_DIR

from shared.parse_main import parse_article, parse_all

sys.path.append(FILE_DIR)


def install(package):
    subprocess.call([sys.executable, "-m", "pip", "install", package])


if __name__ == '__main__':
    install('lxml')
    install('requests')

import sys

def create_debug_dir():
    directory = os.path.join(FILE_DIR, 'debug')

    if not os.path.exists(directory):
        os.makedirs(directory)

    return directory

# def main():
#     parser = argparse.ArgumentParser(prog='NAS Archival', description='Parse URL to NAS .pdf')
#     parser.add_argument('--url', dest='url',
#                         help='url of article REQUIRED 1')
#     parser.add_argument('--year', dest='year', type=int,
#                         help='year of <category> articles REQUIRED 2')
#     parser.add_argument('--category', dest='category', type=int, choices=URL_PARAM_CATEGORY,
#                         help='category of articles REQUIRED 2')
#     parser.add_argument('--debug', dest='debug')
#     args = vars(parser.parse_args())
#
#     if args['year'] is not None and args['category'] is not None:
#         load_logo()
#         listbyyear(category=args['category'], year=args['year'])
#     elif args['url'] is not None:
#         load_logo()
#         debug_directory = create_debug_dir()
#         parse_article(
#             url=args['url'],
#             directory=debug_directory,
#             debug=True)
#     elif args['debug'] is not None:
#         load_logo()
#         debug_directory = create_debug_dir()
#         docx_test()
#     else:
#         parser.print_help()

def main():
    debug_directory = create_debug_dir()
    parse_all(directory=debug_directory)

main()
