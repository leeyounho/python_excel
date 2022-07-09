import pandas as pd
import xlwings as xw

pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 10)

def table_query(book_name, sheet_name):
    # Excel read 시 창이 열리지 않음
    # app = xw.App(visible=False)

    # book, sheet 읽기
    xb = xw.Book(book_name)
    xs = xb.sheets[sheet_name]

    # 'A1' cell 기준으로 table 불러오기
    df1 = xs['A1'].expand().options(pd.DataFrame, header=1, index=False, expand='table').value
    # print(df1)

    # df1 = df1.drop('INSERT_QUERY', axis=1)

    # 열 데이터 초기화
    print(df1)
    col_name = 'INSERT_QUERY'
    col_index = df1.columns.get_loc(col_name)
    df1 = df1.drop(df1.columns[col_index], axis=1)
    df1.insert(col_index, col_name, '')

    col_name = 'ROLLBACK_QUERY'
    col_index = df1.columns.get_loc(col_name)
    df1 = df1.drop(df1.columns[col_index], axis=1)
    df1.insert(col_index, col_name, '')

    df1 = df1.sort_values(by=['TABLE_NAME', 'COLUMN_ID'], ascending=True)

    # Insert Query 생성
    # TODO

    # Rollback query 생성
    # TABLE_NAME 열의 중복제거하여 데이터 삽입
    df1['ROLLBACK_QUERY'] = 'DROP TABLE ' + df1['TABLE_NAME'].drop_duplicates() + ' CASCADE CONSTRAINTS;'

    # Table 저장
    xs['A1'].expand().options(pd.DataFrame, index=False).value = df1


table_query('TC_HELPER.xlsx', 'TABLE')

# purge recycle bin 수행 필요.
