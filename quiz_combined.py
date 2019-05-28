import os

import pandas as pd

from Demographics import *
from ResizeColumn import resize_columns


def quiz_combined(file):
    directory = os.path.dirname(file)
    fileNames = os.listdir(directory)
    files_temp = []
    files_scores_temp = []

    newPath = 'C:/VantagePoint/Full-Quizzes/'
    if not os.path.exists(newPath):
        os.makedirs(newPath)
    if len(fileNames) <= 2:
        raise Exception("Please enter a valid number (Digit not word) For Questions")
        sys.exit("Error")

    # Separate the Scores vs full sheets.
    for f in fileNames:
        if f.endswith(".xlsx"):
                files_temp.append(directory + "/" + f)


    # Converts the excel sheets into a readable and PARSEABLE form.
    quizzes = [pd.ExcelFile(name) for name in files_temp]

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None, na_filter = False) for x in quizzes]
    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first

    # This will keep track of filling the first part of the data frame
    counter = 0
    column = 5
    for index,df in enumerate(frames):
        topic = df.iloc[2,1]
        print(topic)
        df = df.iloc[7:]
        df = df[~df[3].str.startswith('Aver')]
        df[4] = df[4].replace({'Score': topic})
        if counter == 0:
            data = pd.DataFrame(df)
            counter += 1
        if counter == 1:
            data[column] = df[4]
            column+= 1

    data = data.reset_index(drop=True)
    new_header = data.iloc[0]  # grab the first row for the header
    data = data[1:]  # take the data less the header row
    data.columns = new_header  # set the header row as the df header
    data.drop(data.tail(1).index, inplace=True) # drops last row which is empty

    # Make sure to do NA_FILTER = FALSE. This will make sure any and all N/A stays N/A rather than becoming an empty
    # cell.
    # Combines the excel sheets
    data.to_excel(newPath + "Full Quizzes.xlsx", header=True, index=False, na_rep='NA')

    # Resize the columns for the two new excel sheets.
    resize_columns(newPath + "Full Quizzes.xlsx")

    # Alter the scores only to be sorted by age. After we will resort it to be altered by a custom key for race We
    # had to remove the first row, due to Name Age etc messing up the sorting method, and i couldnt find a work
    # around to skip the first row.
    sort_age = data
    # Create sort race for when we will use it in 5 minutes.( 1 millisecond when running)
    sort_race = sort_age.copy()
    print(sort_race)
    # Sorts all the ages in the databse without having to transfer.
    sort_age.sort_values(by="Age", inplace=True)
    print(sort_age)
    # This will find the amount of rows and subtract 1 so we keep the first row as (:-22) will remove 22 rows
    # starting from the bottom.
    # Append it to sort_age
    # Create a new Intake For age Sort
    sort_age.to_excel(newPath + "Full Quizzes-Scores-AgeSort.xlsx", header=True, index=False, na_rep='NA')
    resize_columns(newPath + "Full Quizzes-Scores-AgeSort.xlsx")

    sort_race[4] = pd.Categorical(sort_race["Race"], ["W", "B", "L", "A", "NA", "O", "N/A"])
    sort_race.sort_values(4, inplace=True)

    sort_race.to_excel(newPath + "Full Quizzes-Scores-RaceSort.xlsx", header=True, index=False, na_rep='NA')
    resize_columns(newPath + "Full Quizzes-Scores-RaceSort.xlsx")
    print()

    # combined_scores.to_excel(directory + "/" + "Full Intakes-Scores.xlsx", header=False, index=False, na_rep='NA')
    # resize_columns(directory + "/" + "Full Intakes-Scores-AgeSort.xlsx")
    demographics(newPath + "Full Quizzes-Scores-RaceSort.xlsx", "Race", )
    demographics_number_groups(newPath + "Full Quizzes-Scores-AgeSort.xlsx", "Age", [29, 40, 55])

    sys.exit(0)