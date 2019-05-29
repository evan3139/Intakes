import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from final_quiz_excel import *
import xlsxwriter
import openpyxl




files = []
directories = []
file_names = []
file_titles = []

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
for i in range(len(GROUPING_NAMES) + 1):
    files.append(askopenfilename())  # show an "Open" dialog box and return the path to the selected file

for i,file in enumerate(files):
    directories.append(os.path.dirname(file))
    file_titles.append(os.listdir(directories[i]))

for directory in directories:
    sheets = os.listdir(directory)
    for f in sheets:
        if f.endswith(".xlsx"):
            file_names.append(directory + "/"  + f)


# Every title (E.G. Medium 1, medium 2, minimum etc)
header = []
groups = []

for fi in file_titles:
    for f in fi:
        file = os.path.basename(f)
        file, sep, tail = file.partition("-")
        if file not in header:
            header.append(file)

        # This is because I need which group is being averaged.
        file, sep, tail = tail.partition("Sort")
        file, sep, tail = file.partition("Scores-")
        groups.append(tail)





newPath = 'C:/VantagePoint/QuizTotals/'
if not os.path.exists(newPath):
    os.makedirs(newPath)

total = 'C:/VantagePoint/QuizTotals/Totals.xlsx'
workbook = xlsxwriter.Workbook(total)
worksheet = workbook.add_worksheet()
workbook.close()
format_totals(total, header)

group_index = 0
header_index = 0

for i,f in enumerate(file_names):
    if i == len(header):
        header_index = 0
    quiz_totals(f,total,header[header_index],groups[i])
    header_index += 1


combined_average(total)

