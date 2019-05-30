import openpyxl
import re
import pandas as pd
from pandas import *

def get_topics(excel_name, topics):
    data = pd.ExcelFile(excel_name)
    df = data.parse(data.sheet_names[0],header=None, index_col=None,na_filter=False)

    trigger_word = "id"
    trigger = 0

    for index, row in df.iterrows():
        for cell in row:
            if re.search(r'\d', cell):
                cell = ''.join([i for i in cell if not i.isdigit()])
            cell = cell.rstrip()
            if cell.lower() not in topics and trigger != 0 and cell != '':
                topics.append(cell.lower())
            if cell.lower() == trigger_word:
                trigger += 1
        break
    return topics
