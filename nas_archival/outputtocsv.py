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


def merge_details_files(root_dir):
    output = []

    for r, d, f in os.walk(root_dir):
        details_files = list(filter(lambda v: v == 'details.txt', f))

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
            if not re.search('INFO BUILD BODY', category):
                print(filename, 'Invalid Header: ', category)
            continue

        content = re.sub(r'"', '""', str(res.group(2)))

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


def jsontocsvstr(json, ignore_fields=[]):
    csv = []

    for h in OUTPUT_HEADERS:
        if h in ignore_fields:
            continue

        try:
            if h == RELATED:
                if RELATED in json:
                    for r in json[RELATED]:
                        csv.extend(jsontocsvstr(r, ignore_fields=[START_HEADER]).split(','))
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


def write_to_csv(filename='', transformed_output=[]):
    with open(filename, 'w') as wf:
        wf.write(','.join(list(map(lambda v: "\"{}\"".format(v), DISPLAY_HEADERS))))
        sorted_list = sorted(transformed_output, key=lambda v: int(v['Filename']) if 'Filename' in v else -1)
        csvstr_list = list(map(jsontocsvstr, sorted_list))
        wf.write('\r\n'.join(csvstr_list))


def main(root_dir, output_csv='outputtocsv_results.csv'):
    try:
        os.remove(os.path.join(FILE_DIR, ERROR_FILENAME))
    except FileNotFoundError:
        pass

    merged_file = merge_details_files(root_dir)
    transformed_output = transform_to_output(merged_file)
    output_filename = os.path.join(FILE_DIR, output_csv)
    write_to_csv(filename=output_filename, transformed_output=transformed_output)


if len(sys.argv) != 2:
    print('Required 1 argument <files root directory>')
    sys.exit(1)

files_root_dir = sys.argv[1].strip()

if not os.path.exists(files_root_dir):
    print('{} Directory does not exists'.format(files_root_dir))
    sys.exit(1)

files_root_dir = os.path.abspath(files_root_dir)

start = time.time()
main(files_root_dir)
end = time.time()
print('Processing Time Taken: {} seconds'.format(end - start))
