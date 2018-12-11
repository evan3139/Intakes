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

    newPath = 'C:/VantagePoint/Full-Intakes/'
    if not os.path.exists(newPath):
        os.makedirs(newPath)
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
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None, na_filter=False) for x in files]
    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # Make sure to do NA_FILTER = FALSE. This will make sure any and all N/A stays N/A rather than becoming an empty
    # cell.
    frames_scores = [x.parse(x.sheet_names[0], header=None, index_col=None, na_filter=False) for x in files_scores]
    frames_scores[1:] = [df[1:] for df in frames_scores[1:]]

    # concatenate them..
    combined = pd.concat(frames)
    combined_scores = pd.concat(frames_scores)

    # Combines the excel sheets
    combined.to_excel(newPath + "Full Intakes.xlsx", header=False, index=False, na_rep='NA')
    combined_scores.to_excel(newPath + "Full Intakes-Scores.xlsx", header=False, index=False, na_rep='NA')

    # Resize the columns for the two new excel sheets.
    resize_columns(newPath + "Full Intakes.xlsx")
    resize_columns(newPath + "Full Intakes-Scores.xlsx")

    # Alter the scores only to be sorted by age. After we will resort it to be altered by a custom key for race We
    # had to remove the first row, due to Name Age etc messing up the sorting method, and i couldnt find a work
    # around to skip the first row.
    sort_age = combined_scores.iloc[1:].copy()
    # Create sort race for when we will use it in 5 minutes.( 1 millisecond when running)
    sort_race = sort_age.copy()
    # Sorts all the ages in the databse without having to transfer.
    sort_age.sort_values(2, inplace=True)
    # This will find the amount of rows and subtract 1 so we keep the first row as (:-22) will remove 22 rows
    # starting from the bottom.
    drop = len(combined_scores) - 1
    top_of_file = combined_scores[:-drop]
    # Append it to sort_age
    sort_age = top_of_file.append(sort_age)
    # Create a new Intake For age Sort
    sort_age.to_excel(newPath + "Full Intakes-Scores-AgeSort.xlsx", header=False, index=False, na_rep='NA')
    resize_columns(newPath + "Full Intakes-Scores-AgeSort.xlsx")

    sort_race[4] = pd.Categorical(sort_race[4], ["W", "B", "L", "A", "NA", "O", "N/A"])
    sort_race.sort_values(4, inplace=True)
    sort_race = top_of_file.append(sort_race)

    sort_race.to_excel(newPath + "Full Intakes-Scores-RaceSort.xlsx", header=False, index=False, na_rep='NA')
    resize_columns(newPath + "Full Intakes-Scores-RaceSort.xlsx")

    # combined_scores.to_excel(directory + "/" + "Full Intakes-Scores.xlsx", header=False, index=False, na_rep='NA')
    # resize_columns(directory + "/" + "Full Intakes-Scores-AgeSort.xlsx")

    sys.exit(0)


# This will read a docx file and put every line into a list.
def read_docx(file):
    contents = []
    doc = Document(file)
    for line in doc.paragraphs:
        line.text = line.text.strip()
        if line.text == "":
            continue
        else:
            contents.append(line.text)
    return contents


# This will create the heading for th excel sheet with only the scores in it and not the total intake.
def define_sheet_scores(worksheet, file):
    header = []
    contents = read_docx(file)

    for index, x in enumerate(contents):
        if ":" in x:
            headers, values = x.split(":")
            if headers.lower() == "name":
                header.append(headers)
            elif "date" in x.lower():
                header.append(headers)
            elif "age" in x.lower():
                header.append(headers)
            elif "gender" in x.lower():
                header.append(headers)
            elif "race" in x.lower():
                header.append(headers)
            elif "bdi" in x.lower():
                header.append(headers)
            elif "ace" in x.lower():
                header.append(headers)
            elif "cage" in x.lower():
                header.append(headers)
            elif "bai" in x.lower():
                header.append(headers)

    for index, x in enumerate(header):
        if is_number(header[index]):
            worksheet.write(0, index, int(x))
        else:
            worksheet.write(0, index, x)


# This will create the excel sheet with all the formatting on it for the total inatke sheet
def define_sheet_full(worksheet, file):
    contents = read_docx(file)
    header = []

    for index, x in enumerate(contents):
        if ":" in x:
            headers, values = x.split(":")
            header.append(headers)
        else:
            header.append(x)

    # This will create the heading of the File for intake
    for index, x in enumerate(header):
        if is_number(header[index]):
            worksheet.write(0, index, int(header[index]))
        else:
            worksheet.write(0, index, header[index])


def full_sheet_fill(worksheet, file, row):
    inputs = read_docx(file)
    data = []

    for index, i in enumerate(inputs):
        if ":" in i:
            header, values = i.split(":")
            data.append(values)

    data = [x.strip() for x in data]
    data[1] = data[1].replace(" ", "")

    for index, x in enumerate(data):
        if is_number(x):
            worksheet.write(row, index, int(x))
        else:
            worksheet.write(row, index, x)


# This will fill the data for the scores only sheet.
def score_sheet_fill(worksheet, file, row):
    data = []
    inputs = read_docx(file)

    for index, i in enumerate(inputs):
        if ":" in i:
            useless, values = i.split(":")
            if useless.lower() == "name":
                data.append(values)
            elif "date" in i.lower():
                data.append(values)
            elif "age" in i.lower():
                data.append(values)
            elif "gender" in i.lower():
                data.append(values)
            elif "race" in i.lower():
                data.append(values)
            elif "bdi" in i.lower():
                data.append(values)
            elif "ace" in i.lower():
                data.append(values)
            elif "cage" in i.lower():
                data.append(values)
            elif "bai" in i.lower():
                data.append(values)

    data = [x.strip() for x in data]

    for index, x in enumerate(data):
        if is_number(x):
            worksheet.write(row, index, int(x))
        else:
            worksheet.write(row, index, x)


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
        doc.add_paragraph(
            str(genders[index]) + ',' + str(ages[index]) + "," + str(races[index]) + "," + str(names[index]) + ":")
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

row = 1

# Make all the file names
excel_name = os.path.join("C:/VantagePoint/Intake/" + file_title + ".xlsx")
docx_name = os.path.join("C:/VantagePoint/Quiz-Template/" + file_title + "-QuizTemplate.docx")
scores_name = os.path.join("C:/VantagePoint/Intake/" + file_title + "-Scores.xlsx")

# Create the two Workbooks here
workbook = xlsxwriter.Workbook(excel_name)
worksheet = workbook.add_worksheet()
workbook_scores = xlsxwriter.Workbook(scores_name)
fileNames = os.listdir(directory)
worksheet_scores = workbook_scores.add_worksheet()

# calls the define sheet method which creates the heading of the sheet
define_sheet_full(worksheet, file)
define_sheet_scores(worksheet_scores, file)

for files in fileNames:
    if ".docx" in files:
        filename = directory + "/" + files
        score_sheet_fill(worksheet_scores, workbook_scores, filename, row)
        full_sheet_fill(worksheet, workbook, filename, row)
        row = row + 1
    else:
        continue
workbook.close()
workbook_scores.close()
resize_columns(excel_name)
resize_columns(scores_name)
create_docx_template(excel_name, docx_name, file_title)
