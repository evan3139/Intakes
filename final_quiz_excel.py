import openpyxl
import math
from IsNumber import *

from GLOBALS import *


def combined_average(excel_totals):
    print()

def quiz_totals(excel_data,excel_totals, facility, group):

    read_wb = openpyxl.load_workbook(excel_data,data_only=True)
    read_ws = read_wb.active
    wb = openpyxl.load_workbook(excel_totals,data_only=True)
    ws = wb.active

    groupings = []
    trigger = 0

    for x,col in enumerate(read_ws.columns):
        if col[0].lower() == group.lower():
            for i, cell in enumerate(col):
                if not cell.value:
                    # i is acting as a decrement here since getting a cell starts at 1
                    groupings.append(read_ws.cell(row= x+1, column=i))
            break

    sum = [0] * (len(groupings) + 1)
    divisor = [0] * (len(groupings) + 1)
    rows = [None] * (len(groupings) + 1)
    averages = []

    for col in ws.columns:
        # this will start iterating only after it hits the ID section
        if trigger == 1:
            idx = 0
            for index, cell in enumerate(col):
                if not cell.value:
                    print(idx)
                    rows[idx] = index
                    idx += 1
                elif is_number(cell.value):
                    sum[idx] += cell.value
                    divisor[idx] += 1
        if col[0].value == "ID":
            trigger = 1

    for index, i in enumerate(sum):
        if divisor[index] == 0:
            averages.append(0)
        else:
            averages.append(math.floor(i / divisor[index]))







def format_totals(excel_totals, header):
    wb = openpyxl.load_workbook(excel_totals, data_only=True)
    ws = wb.active

    ws.cell(row=ws.max_row, column=ws.max_column).value = "Ages"

    for group in header:
        ws.cell(row=ws.max_row, column=ws.max_column + 1).value = group

    # skip a column for neatness and put all combined
    ws.cell(row=ws.max_row, column=ws.max_column + 1).value = "All"

    for i, age in enumerate(AGES):
        if i == len(AGES) - 1:
            ws.cell(row=ws.max_row + 1, column=1).value = (str(age) + "+")
        else:
            ws.cell(row=ws.max_row + 1, column=1).value = (str(age) + "-" + str((AGES[i + 1] - 1)))

    group_name = 0

    for group in GROUPINGS:
        # Create the Header
        ws.cell(row=ws.max_row + 2, column=1).value = GROUPING_NAMES[group_name]
        group_name += 1

        for segment in group:
            ws.cell(row=ws.max_row + 1, column=1).value = segment

    wb.save(excel_totals)
