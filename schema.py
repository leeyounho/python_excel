import pandas as pd
import xlwings as xw


def connection_test_to_source_db():
    return


def source_db_load_to_excel():
    return


def connection_test_to_target_db():
    return


def make_table_query(book_name, sheet_name):
    # do not open app
    app = xw.App(visible=False)

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
    app.kill()


def make_index_query():
    return


def make_sequence_query():
    return


def make_procedure_query():
    return


def make_view_query():
    return


def run_table_query():  # TODO rollback도 같은 함수로 구현
    return


def run_index_query():  # TODO rollback도 같은 함수로 구현
    return


def run_sequence_query():  # TODO rollback도 같은 함수로 구현
    return


def run_procedure_query():  # TODO rollback도 같은 함수로 구현
    return


def run_view_query():  # TODO rollback도 같은 함수로 구현
    return


pd.set_option('display.width', 1000)
pd.set_option('display.max_columns', 20)

make_table_query('TC_HELPER.xlsx', 'TABLE')

# purge recycle bin 수행 필요.
