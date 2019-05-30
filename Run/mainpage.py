import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import sys
import openpyxl
import xlrd
import xlsxwriter
from docx import Document
from Demographics import *
from ResizeColumn import *
from ScoreInput import fill_sheets
from CombineSheets import combine_all_sheets
from Header import create_header
from WordTemplate import *


Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
file = askopenfilename()  # show an "Open" dialog box and return the path to the selected file

# If the file is an excel sheet, it will try to combine all of the excel sheets in the folder together.
if file.endswith('.xlsx'):
    combine_all_sheets(file)

# This will open the document, allowing it to be read.
directory = os.path.dirname(file)
file_title = os.path.basename(directory)

# Create the Path needed
newPath = 'C:/VantagePoint/Intake'
if not os.path.exists(newPath):
    os.makedirs(newPath)
newPath = 'C:/VantagePoint/Quiz-Template'
if not os.path.exists(newPath):
    os.makedirs(newPath)

# Make all the file names
excel_name = os.path.join("C:/VantagePoint/Intake/" + file_title + ".xlsx")
docx_name = os.path.join("C:/VantagePoint/Quiz-Template/" + file_title + "-QuizTemplate.docx")
scores_name = os.path.join("C:/VantagePoint/Intake/" + file_title + "-Scores.xlsx")

# Create the two Workbooks here
workbook = xlsxwriter.Workbook(excel_name)
worksheet = workbook.add_worksheet()
workbook_scores = xlsxwriter.Workbook(scores_name)
sheet = workbook_scores.add_worksheet()

# calls the define sheet method which creates the heading of the sheet
create_header(worksheet, sheet, file)

row = 1

fileNames = os.listdir(directory)
for files in fileNames:
    if ".docx" in files:
        filename = directory + "/" + files
        fill_sheets(worksheet, sheet, filename, row)
        row = row + 1
    else:
        continue
workbook.close()
workbook_scores.close()
resize_columns(excel_name)
resize_columns(scores_name)
create_docx_template(excel_name, docx_name, file_title)
