import xlwings as xw
import pandas as pd
from tkinter import *
from tkinter.filedialog import asksaveasfilename
import common as com
import csv
import numpy as np


def save_dataframe_to_file(df, filename, extension):
    if df.empty:
        print('Dataframe is empty')
        return
    save_file_name = asksaveasfilename(initialfile=filename + extension, defaultextension=extension, filetypes=[("All Files", "*.*"), ("SQL Files", "*" + extension)])
    np.savetxt(save_file_name, df.dropna().values, fmt="%s", newline='\n\n')


def save_dataframe_to_file_column_by_column(df, filename, extension):
    if df.empty:
        print('Dataframe is empty')
        return
    save_file_path = asksaveasfilename(initialfile=filename + extension, defaultextension=extension, filetypes=[("All Files", "*.*"), ("SQL Files", "*" + extension)])
    save_file_name = save_file_path.split('/')[-1]
    for idx, column in enumerate(df, 1):
        np.savetxt(f'{idx:02d}' + '_' + save_file_name + extension, df[column].dropna().values, fmt="%s", newline='\n\n')


if __name__ == '__main__':
    xw.Book("TC_HELPER.xlsx").set_mock_caller()

    xb = xw.Book.caller()
    xs = xb.sheets.active
    # xb.sheets['INDEX'].select()
    # print(xs.name)

    # number = 12
    # print('%04d' % (number))
    # print('{:04d}'.format(number))
    # print(f'{number:02d}')

    df1 = com.get_dataframe_from_cell('A1', xs, expand='down')
    print(len(df1.index))
    print(len(df1.columns))
    print(df1.iloc[:, 7].dropna())
    # print()
    save_dataframe_to_file_column_by_column(df1, 'test', '.sql')
