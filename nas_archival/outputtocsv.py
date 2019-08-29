import os
import re
import sys
import time
import traceback

from shared.constants import FILE_DIR
from shared.writers import write_error

RUN_DETAILED = 'detailed'
RUN_SUMMARY = 'summary'
RUN_OPTIONS = [RUN_DETAILED, RUN_SUMMARY]

RELATED_ARTICLE_TYPES = ['FS', 'S']

RELATED_HEADER = 'Related'
SAVE_PATH_HEADER = 'Save Path'
SAVE_FILENAME_HEADER = 'Save Filename'
ERROR_FILENAME = 'outputtocsv_errors.txt'
ARTICLE_TYPE_HEADER = 'Article Type'
URL_HEADER = 'URL'
TITLE_HEADER = 'Title'
DATETIME_HEADER = 'DateTime'
FILENAME_HEADER = 'Filename'
FILENAME_PREFIX_HEADER = 'Filename Prefix'
IS_POSSIBLE_DUP_HEADER = 'Is Possible Duplicate'
PDF_FILES_HEADER = 'PDF Files'
START_HEADER = SAVE_PATH_HEADER
END_ARTICLE_HEADER = SAVE_FILENAME_HEADER
IS_DUPLICATE_HEADER = 'Is Duplicate'
COMBINER = '~~~~~~~~~~~~~~~~~~~~~~~~~'

HEADERS = [PDF_FILES_HEADER, SAVE_PATH_HEADER, ARTICLE_TYPE_HEADER, URL_HEADER, TITLE_HEADER, DATETIME_HEADER, 'Body', 'Images', 'Others Text', 'Others Link',
           FILENAME_HEADER, FILENAME_PREFIX_HEADER, SAVE_FILENAME_HEADER, RELATED_HEADER]
OUTPUT_HEADERS = [SAVE_PATH_HEADER, ARTICLE_TYPE_HEADER, URL_HEADER, TITLE_HEADER, DATETIME_HEADER,
                  FILENAME_HEADER, FILENAME_PREFIX_HEADER, SAVE_FILENAME_HEADER, RELATED_HEADER]
DISPLAY_HEADERS = OUTPUT_HEADERS[:-1] \
                  + list(map(lambda v: '{} 1 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 2 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 3 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 4 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 5 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 6 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 7 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 8 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 9 {}'.format(RELATED_HEADER, v), OUTPUT_HEADERS[1:-1]))
SUMMARY_DISPLAY_HEADERS = ['Type of record* (S=Speech, PR=Press Release, FS=Fact Sheet)',
                           'Title of Speech/Press Release/Fact Sheet',
                           'Date of Press Release or Speech or FS dd mmm yyyy',
                           'NR, FS and Speeches (Filename) MINDEF_yyyymmddNNN.pdf',
                           'Related Resources to Record (Filename) MINDEF_yyyymmddNNN.pdf',
                           'Related Resources in Folder to Record (Filename) MINDEF_yyyymmddNNN.pdf',
                           IS_DUPLICATE_HEADER]
SUMMARY_OUTPUT_HEADERS = [ARTICLE_TYPE_HEADER,
                           TITLE_HEADER,
                           DATETIME_HEADER,
                           SAVE_FILENAME_HEADER,
                           RELATED_HEADER,
                           PDF_FILES_HEADER,
                           IS_DUPLICATE_HEADER]


def merge_details_files(root_dir):
    output = []

    for r, d, f in os.walk(root_dir):
        details_files = list(filter(lambda v: v == 'debug.txt', f))

        for df in details_files:
            full_df = os.path.join(r, df)
            parent_dir = os.path.dirname(full_df)
            pdffiles = list(filter(lambda v: v.endswith('.pdf'), os.listdir(parent_dir)))
            pdffiles.sort()
            pdffiles_str = COMBINER.join(pdffiles)

            with open(full_df, 'r') as f:
                foutput = []
                foutput.append('{}: {}'.format(START_HEADER, full_df))
                foutput.append('{}: {}'.format(PDF_FILES_HEADER, pdffiles_str))
                foutput = foutput + f.readlines()

                if 'news-releases' in full_df:
                    output = foutput + output
                else:
                    output = output + foutput

    return output


def transform_to_output(merged_file):
    is_more_res = False
    category_content_re = re.compile(r'^([^:]+):[ ]*(.+)$')
    filename = ''
    json_obj = dict()
    output = []
    dup_related_checker = dict()

    for line in merged_file:

        if line is None or len(line) == 0:
            continue

        res = re.match(category_content_re, line)

        if not res:
            continue

        category = res.group(1)

        if category not in HEADERS:
            if not re.search('INFO BUILD BODY|https|Fact Sheet|AFTER CHILDREN - COPY PARA & RUN', category):
                print(filename, 'Invalid Header: ', category)
            continue

        content = re.sub(r'"', '""', str(res.group(2)))

        if category == ARTICLE_TYPE_HEADER:
            if '001' in content:
                content = 'NR'
            elif '002' in content:
                content = 'S'
            elif '003' in content:
                content = 'FS'
            else:
                content = 'UNKNOWN'
        elif category == SAVE_FILENAME_HEADER:
            content = re.sub(r'[.]docx', '.pdf', content)

            if not is_more_res:
                if PDF_FILES_HEADER in json_obj:
                    json_obj[PDF_FILES_HEADER] = re.sub(content + '\r?\n?', '', json_obj[PDF_FILES_HEADER])
        elif category == PDF_FILES_HEADER:
            content = re.sub(COMBINER, '\r\n', content)

        if category == START_HEADER:
            dup_checker = dict()
            output.append(json_obj)
            json_obj = dict()
            json_obj[RELATED_HEADER] = []
            is_more_res = False
            filename = content

        if is_more_res:
            if category == URL_HEADER:
                json_obj[RELATED_HEADER].append(dict())

            json_obj[RELATED_HEADER][-1][category] = content

            if category == TITLE_HEADER:
                if content in dup_checker:
                    json_obj[RELATED_HEADER][-1][IS_POSSIBLE_DUP_HEADER] = True
                else:
                    dup_checker[content] = True

            zz = json_obj[RELATED_HEADER][-1]
            if ARTICLE_TYPE_HEADER in zz and TITLE_HEADER in zz and DATETIME_HEADER in zz:
                articletype = zz[ARTICLE_TYPE_HEADER]
                datetime = zz[DATETIME_HEADER]
                title = zz[TITLE_HEADER]

                if articletype not in dup_related_checker:
                    dup_related_checker[articletype] = dict()
                if datetime not in dup_related_checker[articletype]:
                    dup_related_checker[articletype][datetime] = dict()
                dup_related_checker[articletype][datetime][title] = True
        else:
            json_obj[category] = content

            if category == TITLE_HEADER:
                dup_checker[content] = True

            zz = json_obj
            if ARTICLE_TYPE_HEADER in zz and TITLE_HEADER in zz and DATETIME_HEADER in zz:
                articletype = zz[ARTICLE_TYPE_HEADER]
                datetime = zz[DATETIME_HEADER]
                title = zz[TITLE_HEADER]

                if articletype not in dup_related_checker:
                    dup_related_checker[articletype] = dict()
                if datetime not in dup_related_checker[articletype]:
                    dup_related_checker[articletype][datetime] = dict()
                json_obj[IS_DUPLICATE_HEADER] = 'Yes' if title in dup_related_checker[articletype][datetime] else 'No'

        if not is_more_res and category == END_ARTICLE_HEADER:
            is_more_res = True

    output.append(json_obj)

    return output


def jsontocsvstr(json, req_headers=OUTPUT_HEADERS, ignore_fields=[], is_expand_related=True):
    csv = []

    for h in req_headers:
        if h in ignore_fields:
            continue

        try:
            if h == RELATED_HEADER:
                if RELATED_HEADER in json:
                    if is_expand_related:
                        for r in json[RELATED_HEADER]:
                            csv.extend(jsontocsvstr(r, req_headers=req_headers, ignore_fields=[START_HEADER], is_expand_related=is_expand_related).split(','))
                    else:
                        csv.append('"'+'\r\n'.join([r[SAVE_FILENAME_HEADER] + (': POSSIBLE DUPLICATE' if IS_POSSIBLE_DUP_HEADER in r else '') for r in json[RELATED_HEADER]])+'"')

            else:
                strval = str(json[h])
                csv.append("\"{}\"".format(strval))
        except KeyError:
            write_error(directory=FILE_DIR, filename=ERROR_FILENAME,
                        error='KeyError: Header - {}, JSON - {}, Traceback - {}'.format(h, json,
                                                                                        traceback.format_exc()))
            csv.append('')
        except TypeError:
            write_error(directory=FILE_DIR, filename=ERROR_FILENAME,
                        error='TypeError: Header - {}, JSON - {}, Traceback - {}'.format(h, json,
                                                                                         traceback.format_exc()))
            csv.append('')

    return ','.join(csv)


def write_to_csv(filename='', transformed_output=[], display_headers=DISPLAY_HEADERS, req_headers=OUTPUT_HEADERS, is_expand_related=True):
    with open(filename, 'w') as wf:
        wf.write(','.join(list(map(lambda v: "\"{}\"".format(v), display_headers))))
        sorted_list = sorted(transformed_output, key=lambda v: int(v['Filename']) if 'Filename' in v else -1)
        csvstr_list = []

        for l in sorted_list:
            csvstr_list.append(jsontocsvstr(l, req_headers=req_headers, is_expand_related=is_expand_related))

        wf.write('\r\n'.join(csvstr_list))


def main(root_dir, category, output_csv='outputtocsv_results.csv'):
    try:
        os.remove(os.path.join(FILE_DIR, ERROR_FILENAME))
    except FileNotFoundError:
        pass

    merged_file = merge_details_files(root_dir)
    transformed_output = transform_to_output(merged_file)
    output_filename = os.path.join(FILE_DIR, category+'_'+output_csv)

    if category == 'detailed':
        write_to_csv(filename=output_filename, transformed_output=transformed_output)
    elif category == 'summary':
        write_to_csv(filename=output_filename, transformed_output=transformed_output,
                     display_headers=SUMMARY_DISPLAY_HEADERS, req_headers=SUMMARY_OUTPUT_HEADERS, is_expand_related=False)


if len(sys.argv) != 3:
    print('Required 2 arguments <files root directory> <detailed/summary>')
    sys.exit(1)

files_root_dir = sys.argv[1].strip()
output_category = sys.argv[2].strip().lower()

if not os.path.exists(files_root_dir):
    print('{} Directory does not exists'.format(files_root_dir))
    sys.exit(1)

if output_category not in RUN_OPTIONS:
    print('{} Category not in [detailed or summary]'.format(output_category))
    sys.exit(1)

files_root_dir = os.path.abspath(files_root_dir)

start = time.time()
main(files_root_dir, output_category)
end = time.time()
print('Processing Time Taken: {} seconds'.format(end - start))
