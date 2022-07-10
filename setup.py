import pandas as pd
import xlwings as xw


def table_query(book_name, sheet_name):
    # do not open app
    app = xw.App(visible=False)

    # load book & sheet
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    # load table & convert number to string
    df1 = xs['A1'].expand().options(pd.DataFrame, header=1, index=False, expand='table',
                                    numbers=lambda x: str(int(x))).value

    # Initialize column
    col_name = 'INSERT_QUERY'
    col_index = df1.columns.get_loc(col_name)
    df1 = df1.drop(df1.columns[col_index], axis=1)
    df1.insert(col_index, col_name, '')

    col_name = 'ROLLBACK_QUERY'
    col_index = df1.columns.get_loc(col_name)
    df1 = df1.drop(df1.columns[col_index], axis=1)
    df1.insert(col_index, col_name, '')

    # Sort
    df1 = df1.sort_values(by=['TABLE_NAME', 'COLUMN_ID'], ascending=True)

    # Add min, max column_id
    df1['MAX_COLUMN_ID'] = df1.groupby('TABLE_NAME')['COLUMN_ID'].transform('max')
    df1['MIN_COLUMN_ID'] = df1.groupby('TABLE_NAME')['COLUMN_ID'].transform('min')

    # print(df1)

    # Create Table Query
    for index, row in df1.iterrows():
        temp = ''

        # 첫번째 COLUMN_ID면 CREATE TABLE 구문 추가
        if row['COLUMN_ID'] == row['MIN_COLUMN_ID']:
            temp += 'CREATE TABLE ' + row['TABLE_NAME'] + '\n(\n'

        # add column
        temp += row['COLUMN_NAME'] + ' ' + row['DATA_TYPE'] + ' ' + '(' + str(row['DATA_LENGTH']) + ')'

        # add not null
        temp += (' NOT NULL' if row['NULLABLE'] == 'N' else '')

        # add default value
        if row['DATA_DEFAULT']:
            if row['DATA_DEFAULT'] == 'SYSDATE' or row['DATA_DEFAULT'] == 'SYSTIMESTAMP' or row['DATA_DEFAULT'].isdigit():
                temp += ' DEFAULT ' + row['DATA_DEFAULT']
            else:
                temp += ' DEFAULT ' + "'" + row['DATA_DEFAULT'] + "'"

        # add comma
        if row['COLUMN_ID'] != row['MAX_COLUMN_ID']:
            temp += ','
        else:
            temp += '\n);'

        row['INSERT_QUERY'] = temp

    # drop temp column
    df1 = df1.drop(df1[['MAX_COLUMN_ID', 'MIN_COLUMN_ID']], axis=1)

    print(df1)
    # Table Index TODO

    # Drop Table Query
    df1['ROLLBACK_QUERY'] = 'DROP TABLE ' + df1['TABLE_NAME'].drop_duplicates() + ' CASCADE CONSTRAINTS;'

    # Table 저장
    xs['A1'].expand().options(pd.DataFrame, index=False).value = df1

    app.kill()

pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 10)

table_query('TC_HELPER.xlsx', 'TABLE')

# purge recycle bin 수행 필요.
