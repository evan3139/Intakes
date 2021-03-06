import os

import pandas as pd

from Demographics import *
from ResizeColumn import resize_columns
from GLOBALS import *


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
    frames = [x.parse(x.sheet_names[0], header=None, index_col=None, na_filter = False) for x in files]
    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # Make sure to do NA_FILTER = FALSE. This will make sure any and all N/A stays N/A rather than becoming an empty
    # cell.
    frames_scores = [x.parse(x.sheet_names[0], header=None, index_col=None, na_filter = False) for x in files_scores]
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

    sort_race[4] = pd.Categorical(sort_race[4], RACES)
    sort_race.sort_values(4, inplace=True)
    sort_race = top_of_file.append(sort_race)

    sort_race.to_excel(newPath + "Full Intakes-Scores-RaceSort.xlsx", header=False, index=False, na_rep='NA')
    resize_columns(newPath + "Full Intakes-Scores-RaceSort.xlsx")

    # combined_scores.to_excel(directory + "/" + "Full Intakes-Scores.xlsx", header=False, index=False, na_rep='NA')
    # resize_columns(directory + "/" + "Full Intakes-Scores-AgeSort.xlsx")
    demographics(newPath + "Full Intakes-Scores-RaceSort.xlsx","Race",)
    demographics_number_groups(newPath + "Full Intakes-Scores-AgeSort.xlsx","Age",AGES)

    sys.exit(0)