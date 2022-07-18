import xlwings as xw
import pandas as pd
from tkinter import *
from tkinter.filedialog import asksaveasfile
import common as com
import csv
import numpy as np


def save_dataframe_to_file(df, filename, extension):
    save_file = asksaveasfile(initialfile=filename + extension, defaultextension=extension, filetypes=[("All Files", "*.*"), ("SQL Files", "*" + extension)])
    np.savetxt(save_file.name, df.dropna().values, fmt="%s", newline='\n\n')


if __name__ == '__main__':
    xw.Book("TC_HELPER.xlsx").set_mock_caller()

    xb = xw.Book.caller()
    xs = xb.sheets.active
    xb.sheets['INDEX'].select()
    # print(xs.name)

    # df1 = com.get_dataframe_from_cell('A1', xs, expand='down')
    # print(df1.iloc[:,7].dropna())
    # print()
    # save_dataframe_to_file(df1.iloc[:,7], 'test', '.sql')
