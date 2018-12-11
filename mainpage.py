import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import sys
import openpyxl
import pandas as pandas
import pandas as pd
import xlrd
import xlsxwriter
from docx import Document


def is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


def combine_all_sheets(file):
    directory = os.path.dirname(file)
    fileNames = os.listdir(directory)
    files_temp = []
    files_scores_temp = []
    if len(fileNames) <= 2:
        raise Exception("Please enter a valid number (Digit not word) For Questions")
        sys.exit("Error")

    # Separate the Scores vs full sheets.
    for f in fileNames:
        if f.endswith(".xlsx") and "Full" not in f:
            if "Scores" not in f:
                files_temp.append(directory + "/" + f)
            else:
                files_scores_temp.append(directory + "/" + f)

    # Converts the excel sheets into a readable and PARSEABLE form.
    files = [pd.ExcelFile(name) for name in files_temp]
    files_scores = [pd.ExcelFile(name) for name in files_scores_temp]

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None) for x in files]
    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    frames_scores = [x.parse(x.sheet_names[0], header=None, index_col=None) for x in files_scores]
    frames_scores[1:] = [df[1:] for df in frames_scores[1:]]

    # concatenate them..
    combined = pd.concat(frames)
    combined_scores = pd.concat(frames_scores)

    # Combines the excel sheets
    combined.to_excel(directory + "/" + "Full Intakes.xlsx", header=False, index=False)
    combined_scores.to_excel(directory + "/" + "Full Intakes-Scores.xlsx", header=False, index=False)

    # Resize the columns for the two new excel sheets.
    resize_columns(directory + "/" + "Full Intakes.xlsx")
    resize_columns(directory + "/" + "Full Intakes-Scores.xlsx")

    sys.exit(0)


# This will create the excel sheet with all the formatting on it
def define_sheet(worksheet, worksheet_scores, file):
    contents = []
    header = []
    content = []
    score_header = []

    doc = Document(file)

    for line in doc.paragraphs:
        line.text = line.text.strip()
        if line.text == "":
            continue
        else:
            contents.append(line.text)

    for index, x in enumerate(contents):
        if ":" in x:
            headers, values = x.split(":")
            header.append(headers)
            content.append(values)
            if headers.lower() == "name":
                score_header.append(headers)
            elif "date" in x.lower():
                score_header.append(headers)
            elif "age" in x.lower():
                score_header.append(headers)
            elif "gender" in x.lower():
                score_header.append(headers)
            elif "race" in x.lower():
                score_header.append(headers)
            elif "bdi" in x.lower():
                score_header.append(headers)
            elif "ace" in x.lower():
                score_header.append(headers)
            elif "cage" in x.lower():
                score_header.append(headers)
            elif "bai" in x.lower():
                score_header.append(headers)
        else:
            header.append(x)

    # This will create the heading of the File for intake
    for index, x in enumerate(header):
        if is_number(header[index]):
            worksheet.write(0, index, int(header[index]))
        else:
            worksheet.write(0, index, header[index])

    for index, x in enumerate(score_header):
        if is_number(score_header[index]):
            worksheet_scores.write(0, index, int(x))
        else:
            worksheet_scores.write(0, index, x)


# This wi
def score_sheet(worksheet, worksheet_scores, filename, row):
    doc = Document(filename)
    inputs = []
    data = []
    score_data = []

    for line in doc.paragraphs:
        line.text = line.text.strip()
        if line.text == "":
            continue
        else:
            inputs.append(line.text)

    for index, i in enumerate(inputs):
        if ":" in i:
            useless, values = i.split(":")
            data.append(values)
            if useless.lower() == "name":
                score_data.append(values)
            elif "date" in i.lower():
                score_data.append(values)
            elif "age" in i.lower():
                score_data.append(values)
            elif "gender" in i.lower():
                score_data.append(values)
            elif "race" in i.lower():
                score_data.append(values)
            elif "bdi" in i.lower():
                score_data.append(values)
            elif "ace" in i.lower():
                score_data.append(values)
            elif "cage" in i.lower():
                score_data.append(values)
            elif "bai" in i.lower():
                score_data.append(values)
        else:
            data.append(i)

    data = [x.strip() for x in data]
    score_data = [x.strip() for x in score_data]
    data[1] = data[1].replace(" ", "")

    for index, x in enumerate(data):
        if is_number(x):
            worksheet.write(row, index, int(x))
        else:
            worksheet.write(row, index, x)

    for index, x in enumerate(score_data):
        if is_number(x):
            sheet.write(row, index, int(x))
        else:
            sheet.write(row, index, x)


def resize_columns(excel_name):
    # This reopens the excel file but in the openpyxl library allowing us to alter column lengths
    wb = openpyxl.load_workbook(excel_name)
    worksheet = wb.active

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column  # Get the Column Name Here
        for cell in col:
            try:  # Needed to avoid empty cell errors
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

    wb.save(excel_name)


def create_docx_template(excel_name, docx_name, file_title):
    # Here will create a template in a docx file for the quiz templates
    workbook = xlrd.open_workbook(excel_name)
    worksheet = workbook.sheet_by_index(0)
    date = worksheet.cell_value(0, 0)

    # Creates the needed arrays
    names = []
    ages = []
    genders = []
    races = []

    # Grabs all the names ages genders and races through their columns and puts them in arrays.
    for i in range(worksheet.nrows):
        if i > 0:
            names.append(worksheet.cell_value(i, 1))
            ages.append(int(worksheet.cell_value(i, 5)))
            genders.append(worksheet.cell_value(i, 6))
            races.append(worksheet.cell_value(i, 12))
    doc = Document()

    # Creates the top of the file with stuff that will not change.
    doc.add_paragraph(date + ":")
    doc.add_paragraph("Facilitator:")
    doc.add_paragraph("Topic:")
    doc.add_paragraph("Week:")
    doc.add_paragraph("Group:" + file_title)
    doc.add_paragraph("Questions:")
    doc.add_paragraph("")

    # Fills the file with ID's genders, ages, races in a template format.
    for index, i in enumerate(names):
        doc.add_paragraph(genders[index] + ',' + str(ages[index]) + "," + races[index] + "," + names[index] + ":")
    doc.save(docx_name)


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
define_sheet(worksheet, sheet, file)

row = 1

fileNames = os.listdir(directory)
for files in fileNames:
    if ".docx" in files:
        filename = directory + "/" + files
        score_sheet(worksheet, workbook_scores, filename, row)
        row = row + 1
    else:
        continue
workbook.close()
workbook_scores.close()
resize_columns(excel_name)
resize_columns(scores_name)
create_docx_template(excel_name, docx_name, file_title)
