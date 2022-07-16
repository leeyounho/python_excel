import xlwings as xw
import pandas as pd
from tkinter import *
from tkinter.filedialog import asksaveasfile


def save_dataframe_to_file(df, filename, extension, skip_na=True):
    save_file = asksaveasfile(initialfile=filename + extension, defaultextension=extension, filetypes=[("All Files", "*.*"), ("SQL Files", "*" + extension)])

    # TODO if canceled
    if save_file:
        my_file = open(save_file.name, 'w')
        for data in df.columns:
            if skip_na :
                my_file.write(df[data].dropna().to_string(index=False) + '\n')
            else :
                my_file.write(df[data].to_string(index=False) + '\n')
        my_file.close()
        save_file.close()
        print('save file dialog worked')
    else:
        print('save file dialog was closed')


if __name__ == '__main__':
    xw.Book("TC_HELPER.xlsx").set_mock_caller()

    xb = xw.Book.caller()
    xs = xb.sheets.active
    print(xs.name)

    # df1 = xs['A:A'].value
    # print(df1)
    # save_dataframe_to_file(df1, 'test', '.sql', False)

