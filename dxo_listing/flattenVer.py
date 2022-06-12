import xlsxwriter
import json

CELL_WIDTH = 50
HIGHEST = 'HIGHEST_COLUMN_NUMBER'


def get_first_letters(val):
    words = val.split()
    if len(words) == 1 and len(words[0]) > 1:
        return words[0][0:2]
    first_letters = [word[0] for word in words]
    first_letters_str = "".join(first_letters)
    return first_letters_str


def read_json_list(json_file, wb, ws, prefix, key_val_list,
                   ws_row, ws_col_map, title_format, value_format,
                   parent=None):
    new_ws_row = ws_row
    for item in json_file:
        # add parent values
        if parent is not None:
            for k, val in parent.items():
                key = k.strip()
                if type(val) is not list:
                    # write value to column by key and column map
                    val_str = val.strip()
                    ws.write(ws_col_map[key] + str(new_ws_row), val_str, value_format)
        # add item values
        new_ws_row = read_json_dict(item, wb, ws, prefix, key_val_list,
                                    new_ws_row, ws_col_map, title_format, value_format)
        new_ws_row += 1
        print('list', prefix, new_ws_row)
    return new_ws_row - 1


def read_json_dict(json_file, wb, ws, prefix, key_val_list,
                   ws_row, ws_col_map, title_format, value_format):
    new_ws = ws
    new_ws_row = ws_row
    new_ws_col_map = ws_col_map

    for k, val in json_file.items():
        key = k.strip()
        # write value to column by key and column map
        if type(val) is not list:
            val_str = val.strip()
            new_ws.write(new_ws_col_map[key] + str(new_ws_row), val_str, value_format)

    for k, val in json_file.items():
        key = k.strip()
        if type(val) is list:
            # generate parent prefix for worksheet name
            new_prefix = ''
            is_require_new_worksheet = False
            for j, h in enumerate(key_val_list):
                if h in json_file:
                    is_require_new_worksheet = j == 0
                    new_prefix = prefix
                    if len(new_prefix) > 0:
                        new_prefix = new_prefix + ' - '
                    new_prefix = new_prefix + get_first_letters(json_file[h])
            # write worksheet name to column
            new_ws.write(new_ws_col_map[key] + str(new_ws_row), new_prefix, value_format)
            # create new worksheet using new prefix
            if is_require_new_worksheet:
                new_ws = wb.add_worksheet(new_prefix)
                # add titles to new worksheet, generate column map by new title and column
                i = ord('A')
                new_ws_col_map = dict()
                new_ws_row = 2
            else:
                # add titles to current worksheet, generate column map by new title and column
                i = new_ws_col_map[HIGHEST] + 1
                json_file[k] = new_prefix
            for title in val[0].keys():
                if title not in new_ws_col_map:
                    col = chr(i)
                    new_ws.set_column(i - 1, 0, CELL_WIDTH)
                    new_ws.write(col + '1', title, title_format)
                    new_ws_col_map[title] = col
                    new_ws_col_map[HIGHEST] = i
                    i += 1
            # add other rows or new worksheets
            result = read_json_list(val, wb, new_ws, new_prefix, key_val_list,
                                        new_ws_row, new_ws_col_map, title_format, value_format,
                                        json_file if not is_require_new_worksheet else None)
            if not is_require_new_worksheet:
                new_ws_row = result
            print(new_prefix, new_ws_row)
    return new_ws_row


def run():
    with open('vacancies-positions-listing.json', encoding='utf-8-sig') as f:
        json_file = json.load(f)
    with xlsxwriter.Workbook('vacancies-positions-listing.xlsx') as wb:
        json_file = json_file['jobgroup']
        ws = wb.add_worksheet('Job Groups')
        title_format = wb.add_format()
        title_format.set_bold()
        title_format.set_border()
        value_format = wb.add_format()
        value_format.set_border()
        value_format.set_text_wrap()
        # add job groups titles to worksheet, generate column map by new title and column
        i = ord('A')
        new_ws_col_map = dict()
        for title in json_file[0].keys():
            col = chr(i)
            ws.set_column(i - 1, 0, CELL_WIDTH)
            ws.write(col + '1', title, title_format)
            new_ws_col_map[title] = col
            new_ws_col_map[HIGHEST] = i
            i += 1
        new_ws_row = 2
        # add other worksheets
        read_json_list(json_file, wb, ws, '', ['jobdepartment', 'jobdepartmentpostion'],
                       new_ws_row, new_ws_col_map, title_format, value_format)


if __name__ == '__main__':
    run()
