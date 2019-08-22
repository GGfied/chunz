import os
import sys
import re
import traceback

from shared.constants import FILE_DIR

RELATED = 'Related'
START_HEADER = 'Save Path'
END_ARTICLE_HEADER = 'Save Filename'

if len(sys.argv) != 2:
    print('Required 1 argument <filename>')
    sys.exit(1)

filename = sys.argv[1].strip()
HEADERS = ['Save Path', 'Article Type', 'URL', 'Title', 'DateTime', 'Body', 'Images', 'Others Text', 'Others Link',
           'Filename', 'Filename Prefix', 'Save Filename', RELATED]
OUTPUT_HEADERS = ['Save Path', 'Article Type', 'URL', 'Title', 'DateTime', 'Filename', 'Filename Prefix', 'Save Filename', RELATED]
DISPLAY_HEADERS = OUTPUT_HEADERS[:-1]\
                  + list(map(lambda v: '{} 1 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 2 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 3 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 4 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 5 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 6 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 7 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 8 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))\
                  + list(map(lambda v: '{} 9 {}'.format(RELATED, v), OUTPUT_HEADERS[1:-1]))
output = []

with open(filename, 'r') as rf:
    json_obj = dict()
    is_more_res = False
    restr = re.compile(r'^([^:]+):[ ]*(.+)$')

    while True:
        line = rf.readline()

        if line is None or len(line) == 0:
            break

        res = re.match(restr, line)

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

        if is_more_res:
            if category == 'URL':
                json_obj[RELATED].append(dict())
            json_obj[RELATED][-1][category] = content
        else:
            json_obj[category] = content

        if not is_more_res and category == END_ARTICLE_HEADER:
            is_more_res = True

    output.append(json_obj)

output_filename = os.path.join(FILE_DIR, 'output.csv')


def jsontocsvstr(json, ignore_fields=[]):
    csv = []

    for h in OUTPUT_HEADERS:
        if h in ignore_fields:
            continue

        try:
            if h == RELATED:
                if RELATED in json:
                    for r in json[RELATED]:
                        csv.extend(jsontocsvstr(r, ignore_fields=['Save Path']).split(','))
            else:
                strval = str(json[h])
                csv.append("\"{}\"".format(strval))
        except KeyError:
            traceback.print_exc()
            print('KeyError', 'header', h, 'json', json)
            csv.append('')
        except TypeError:
            traceback.print_exc()
            print('TypeError', 'header', h, 'json', json)
            csv.append('')

    return ','.join(csv)


with open(output_filename, 'w') as wf:
    wf.write(','.join(list(map(lambda v: "\"{}\"".format(v), DISPLAY_HEADERS))))
    sorted_list = sorted(output, key=lambda v: int(v['Filename']) if 'Filename' in v else -1)
    csvstr_list = list(map(jsontocsvstr, sorted_list))
    wf.write('\r\n'.join(csvstr_list))


