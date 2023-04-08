import re
import pandas as pd
import pathlib
import os

def path_system():
    THIS_FOLDER = os.path.dirname(os.path.abspath(__file__)) + "/files"
    my_file = os.path.join(THIS_FOLDER, 'sat.xls')
    return my_file

def get_sheets():
    xls = pd.ExcelFile(path_system())
    sheets = xls.sheet_names
    print(sheets[0])


if __name__ == '__main__':
    get_sheets()