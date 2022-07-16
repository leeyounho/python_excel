import traceback

import pandas as pd
import xlwings as xw

# rgb 색상표 http://www.n2n.pe.kr/lev-1/color.htm
black_color = (0, 0, 0)
gray = (128, 128, 128)

def make_table_query(book_name, sheet_name):
    # do not open app
    app = xw.App(visible=False)

    try:
        # load book & sheet
        xb = xw.Book(book_name)
        xs = xb.sheets[sheet_name]

        # load table & convert number to string
        df1 = xs['A1'].expand().options(pd.DataFrame, header=1, index=False, expand='table',
                                        numbers=lambda x: str(int(x))).value

        # Sort
        df1 = df1.sort_values(by=['TABLE_NAME', 'COLUMN_ID'], ascending=True)

        df2 = pd.DataFrame(columns={'TABLE_NAME', 'COLUMN_ID', 'INSERT_QUERY'})

        for name, group in df1.groupby('TABLE_NAME'):
            max_i = len(group) - 1
            string = 'CREATE TABLE ' + name + '\n(\n'
            for i, (index, row) in enumerate(group.iterrows()):
                # add column
                string += row['COLUMN_NAME'] + ' ' + row['DATA_TYPE'] + ' ' + '(' + row['DATA_LENGTH'] + ')'

                # add not null
                string += (' NOT NULL' if row['NULLABLE'] == 'N' else '')

                # add default value
                if row['DATA_DEFAULT']:
                    if row['DATA_DEFAULT'] == 'SYSDATE' or row['DATA_DEFAULT'] == 'SYSTIMESTAMP' or row['DATA_DEFAULT'].isdigit():
                        string += ' DEFAULT ' + row['DATA_DEFAULT']
                    else:
                        string += ' DEFAULT ' + "'" + row['DATA_DEFAULT'] + "'"

                # add comma
                string += ',\n' if i != max_i else '\n'

            # TODO add tablespace
            string += ');'

            # TODO deprecated
            df2 = df2.append({'TABLE_NAME': name, 'COLUMN_ID': '1', 'INSERT_QUERY': string}, ignore_index=True)

        # INSERT_QUERY Merge
        df1 = df1.drop('INSERT_QUERY', axis=1)
        df1 = pd.merge(df1, df2, how='left', on=['TABLE_NAME', 'COLUMN_ID'])

        # Table Index TODO

        # Drop Table Query
        df1 = df1.drop('ROLLBACK_QUERY', axis=1)
        df1['ROLLBACK_QUERY'] = 'DROP TABLE ' + df1['TABLE_NAME'].drop_duplicates() + ' CASCADE CONSTRAINTS;'

        # Table 저장
        xs['A1'].expand().options(pd.DataFrame, index=False).value = df1

        print(df1)
    except Exception as e:
        app.kill()
        traceback.print_exc()

def delete_range(book_name, sheet_name, range_string):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    xs.range(range_string).delete()  # TODO xlwings range column number
    print(range_string + ' column deleted')


def align_range(book_name, sheet_name, align_string, range_string):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    if align_string == 'LEFT':
        xs.range(range_string).api.HorizontalAlignment = -4108
    elif align_string == 'CENTER':
        xs.range(range_string).api.HorizontalAlignment = -4131
    elif align_string == 'RIGHT':
        xs.range(range_string).api.HorizontalAlignment = -4152
    else:
        xs.range(range_string).api.HorizontalAlignment = -4131


def color_background_range(book_name, sheet_name, range_string, rgb):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    xs.range(range_string).color = rgb


def bold_font_range(book_name, sheet_name, range_string):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    xs.range(range_string).api.Font.Bold = True


def set_border_range(book_name, sheet_name, range_string, line_thickness):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    # api return native object pywin32
    # line_thickness = 1:dotted, 2:straight, 3:bold
    xs.range(range_string).api.Borders.Weight = line_thickness


def set_font_size_range(book_name, sheet_name, range_string, font_size):
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    xs.range(range_string).api.Font.Size = font_size


# pd.set_option('display.width', 1000)
# pd.set_option('display.max_columns', 20)

# make_table_query('TC_HELPER.xlsx', 'TABLE')

# delete_range('TC_HELPER.xlsx', '5', 'C:D')
# set_font_size_range('TC_HELPER.xlsx', '5', 'C6:E9', 12)

# purge recycle bin 수행 필요.

def get_book_name():
    return xw.Book.caller()


def get_book_name():
    return xw.Book.caller().sheets.active

if __name__ == '__main__':
    print('main test')

