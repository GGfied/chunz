import os
import re
import sys
import time
import traceback

from shared.constants import FILE_DIR
from shared.writers import write_error

RELATED = 'Related'
SAVE_PATH = 'Save Path'
START_HEADER = SAVE_PATH
END_ARTICLE_HEADER = 'Save Filename'
ERROR_FILENAME = 'outputtocsv_errors.txt'

HEADERS = [SAVE_PATH, 'Article Type', 'URL', 'Title', 'DateTime', 'Body', 'Images', 'Others Text', 'Others Link',
           'Filename', 'Filename Prefix', 'Save Filename', RELATED]
OUTPUT_HEADERS = ['Save Path', 'Article Type', 'URL', 'Title', 'DateTime', 'Filename', 'Filename Prefix',
                  'Save Filename', RELATED]
DISPLAY_HEADERS = OUTPUT_HEADERS[:-1] \
                  + list(map(lambda v: '{} 1 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 2 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 3 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 4 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 5 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 6 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 7 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 8 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1])) \
                  + list(map(lambda v: '{} 9 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))
SUMMARY_DISPLAY_HEADERS = ['Type of record* (S=Speech, PR=Press Release, FS=Fact Sheet)',
                           'Title of Speech/Press Release/Fact Sheet',
                           'Date of Press Release or Speech or FS dd mmm yyyy',
                           'NR, FS and Speeches (Filename) MINDEF_yyyymmddNNN.pdf',
                           'Related Factsheets to PR (Filename) MINDEF_yyyymmddNNN.pdf']
SUMMARY_OUTPUT_HEADERS = ['Article Type',
                           'Title',
                           'DateTime',
                           'Save Filename',
                           'Related']


def merge_details_files(root_dir):
    output = []

    for r, d, f in os.walk(root_dir):
        details_files = list(filter(lambda v: v == 'debug.txt', f))

        for df in details_files:
            full_df = os.path.join(r, df)

            with open(full_df, 'r') as f:
                output.append('{}: {}'.format(START_HEADER, full_df))
                output = output + f.readlines()

    return output


def transform_to_output(merged_file):
    is_more_res = False
    category_content_re = re.compile(r'^([^:]+):[ ]*(.+)$')
    filename = ''
    json_obj = dict()
    output = []

    for line in merged_file:

        if line is None or len(line) == 0:
            continue

        res = re.match(category_content_re, line)

        if not res:
            continue

        category = res.group(1)

        if category not in HEADERS:
            if not re.search('INFO BUILD BODY|https|Fact Sheet', category):
                print(filename, 'Invalid Header: ', category)
            continue

        content = re.sub(r'"', '""', str(res.group(2)))

        if category == 'Article Type':
            if '001' in content:
                content = 'NR'
            elif '002' in content:
                content = 'S'
            elif '003' in content:
                content = 'FS'
            else:
                content = 'UNKNOWN'

        if category == 'Save Filename':
            content = re.sub(r'[.]docx', '.pdf', content)

        if category == START_HEADER:
            output.append(json_obj)
            json_obj = dict()
            json_obj[RELATED] = []
            is_more_res = False
            filename = content

        if is_more_res:
            if category == 'URL':
                json_obj[RELATED].append(dict())
            json_obj[RELATED][-1][category] = content
        else:
            json_obj[category] = content

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
            if h == RELATED:
                if RELATED in json:
                    if is_expand_related:
                        for r in json[RELATED]:
                            csv.extend(jsontocsvstr(r, req_headers=req_headers, ignore_fields=[START_HEADER], is_expand_related=is_expand_related).split(','))
                    else:
                        csv.append('"'+'\r\n'.join([r['Save Filename'] for r in json[RELATED]])+'"')

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

if output_category not in ['detailed', 'summary']:
    print('{} Category not in [detailed or summary]'.format(output_category))
    sys.exit(1)

files_root_dir = os.path.abspath(files_root_dir)

start = time.time()
main(files_root_dir, output_category)
end = time.time()
print('Processing Time Taken: {} seconds'.format(end - start))
