from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
import xlsxwriter
from docx import Document
from quiz_combined import *
from ReadDocx import *

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
# This will open the document, allowing it to be read.

directory = os.path.dirname(filename)
file_title = os.path.basename(directory)
fileNames = os.listdir(directory)

for f in fileNames:
    if f.endswith(".docx"):
        read_doc(directory + '/' + f,)

if filename.endswith('.xlsx'):
    quiz_combined(filename)

