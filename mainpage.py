from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
from datetime import date
import xlsxwriter
from docx import Document
from docx.shared import Inches
import openpyxl
from os import listdir
from os.path import isfile, join
import xlrd
from openpyxl.worksheet import Worksheet


def is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


# This will create the excel sheet with all the formatting on it
def define_sheet(worksheet, file):
    contents = []
    header = []
    content = []
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
        else:
            header.append(x)

    # This will create the heading of the File for intake
    for index, x in enumerate(header):
        if is_number(header[index]):
            worksheet.write(0, index, int(header[index]))
        else:
            worksheet.write(0, index, header[index])


def alter_sheet(worksheet, filename, row):
    doc = Document(filename)
    inputs = []
    data = []

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
        else:
            data.append(i)

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


def create_docx_template(excel_name, docx_name):
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
    doc.add_paragraph("")

    # Fills the file with ID's genders, ages, races in a template format.
    for index,i in enumerate(names):
        doc.add_paragraph(genders[index] + ',' + str(ages[index]) + "," + races[index] + "," + names[index] + ":")
    doc.save(docx_name)



Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
file = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
# This will open the document, allowing it to be read.
directory = os.path.dirname(file)
file_title = os.path.basename(directory)

# Create the Path needed
newPath = 'C:/Desktop/Intake'
if not os.path.exists(newPath):
    os.makedirs(newPath)
newPath = 'C:/Desktop/Quiz-Template'
if not os.path.exists(newPath):
    os.makedirs(newPath)

excel_name = os.path.join("C:/Desktop/Intake/" + file_title + ".xlsx")
docx_name = os.path.join("C:/Desktop/Quiz-Template/" + file_title + "-QuizTemplate.docx")

workbook = xlsxwriter.Workbook(excel_name)
worksheet = workbook.add_worksheet()
# calls the define sheet method which creates the heading of the sheet
define_sheet(worksheet, file)

row = 1

fileNames = os.listdir(directory)
for files in fileNames:
    if ".docx" in files:
        filename = directory + "/" + files
        alter_sheet(worksheet, filename, row)
        row = row + 1
    else:
        continue
workbook.close()
resize_columns(excel_name)
create_docx_template(excel_name,docx_name)