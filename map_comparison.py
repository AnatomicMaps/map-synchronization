import json
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter.messagebox import showinfo
import pprint
import pandas as pd
import openpyxl as op
import xlsxwriter as xw
import xlsxwriter.exceptions

## Define a function to load JSON files.
def loadJSON():

    file_path = fd.askopenfilename(
        filetypes= (
            ('JSON files', '*.json'),
            ('All files', '*.*')
        )
    )

    file = open(file_path)
    file_data = json.load(file)

    return file_data
##############################

## Define a function to save a file in the JSON format.
def saveFile(data):

    save_path = fd.asksaveasfilename(
        filetypes = (
            ('JSON files', '*.json'),
            ('All files', '*.*')
        ),
        defaultextension = '.json'
    )
    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
##############################

## Define a function to sort a dictionary by the label's alphabetical order.
def sortFn(dict):

    return dict['label']
##############################

## Define a function to input two lists and extract the differences between the two.
def getDifference(list1, list2):
    i = 0
    j = 0
    diff = 0
    differences = []

# Loop comparing the two lists.
    while i < len(list1):
        while j < len(list2):
            if i == len(list1):
                break
            else:
                x = list1[i]
                y = list2[j]

                if x['term'] == y['term']:
                    i += 1
                    diff = 0
                    j = 0
                else:
                    diff += 1
                    j += 1

        if diff > 0:
            differences.append(list1[i])

        i += 1
        diff = 0
        j = 0

# Sort the list by the label's alphabetical order.
    differences.sort(key=sortFn)

    return differences
##############################

## Define a function to input two lists and extract the similarities between the two.
def getSimilar(list1, list2):
    i = 0
    j = 0
    similarities = []

# Loop to extract similarities.
    while i < len(list1):
        while j < len(list2):
            x = list1[i]
            y = list2[j]

            if x['term'] == y['term']:
                similarities.append(list1[i])
                j += 1
            else:
                j += 1

        i += 1
        j = 0

# Sort the list by the label's alphabetical order.
    similarities.sort(key=sortFn)

    return similarities
##############################

## Define a function that inputs two lists, whether the differences or similarities are to be extracted.
## Note that the 'duplicates' argument is whether to remove duplicates within the list or not.
def comparison(task, list1, list2, duplicates=True):

# Need to ensure the desired task is input.
# This will be modified to account for wrong inputs.
    if task == 'differences':
        listWithDuplicates = getDifference(list1, list2)
    elif task == 'similarities':
        listWithDuplicates = getSimilar(list1, list2)

    if duplicates == True:
        listWithoutDuplicates = removeDuplicate(listWithDuplicates)
        return listWithoutDuplicates
    else:
        return listWithDuplicates
##############################

## Define a function to write out to an excel spreadsheet.
## The required arguments are the worksheet name and data.
## Optional arguments are a header, the row and column start point, and the titles of the 3 data columns.
## The 'duplicates' argument is whether the number of occurences of the entry is printed out or not (WIP).
def writeSheet(worksheet, data, header=None, row=0, col=0, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=False):

    ws = wb.add_worksheet(worksheet)

# The header will be printed in the desired start column.
    if header != None:
        ws.write(0, col, header)
        col += 1

# Write out the titles of the data column.
    ws.write(0, col, A1Title)
    ws.write(0, col + 1, B1Title)
    ws.write(0, col + 2, C1Title)

# Write out the title for the count of duplicates, if included.
    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates')

    i = 0
    row = 1

# Loop to write out the data.
    while i < len(data):
        x = data[i]
        id = x['id']
        term = x['term']
        label = x['label']

        ws.write(row, col, label)
        ws.write(row, col + 1, term)
        ws.write(row, col + 2, id)

        if duplicates == True:
            duplicate = x['duplicates']
            ws.write(row, col + 3, duplicate)

        row += 1
        i += 1

    return
##############################

## Define a function to write out to an existing excel spreadsheet.
## The required arguments are the worksheet name, data, and starting row and column.
## The optional arguments are the header name and data titles.
## The duplicates argument is to whether to include the count of duplicates in the data or not.
def addToSheet(worksheet, data, row, col, header=None, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=False):

    ws = wb.get_worksheet_by_name(worksheet)

# Write out the header in the starting column.
    if header != None:
        ws.write(0, col, header)
        col += 1

# Write out the data titles.
    ws.write(0, col, A1Title)
    ws.write(0, col + 1, B1Title)
    ws.write(0, col + 2, C1Title)

# Ensure the data titles in row 0 are not overwritten.
    if row == 0:
        row += 1

# Write out the title for the count of duplicates, if included.
    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates')

    i = 0

# Loop to write out data.
    while i < len(data):
        x = data[i]
        id = x['id']
        term = x['term']
        label = x['label']

        ws.write(row, col, label)
        ws.write(row, col + 1, term)
        ws.write(row, col + 2, id)

        if duplicates == True:
            duplicate = x['duplicates']
            ws.write(row, col + 3, duplicate)

        row += 1
        i += 1

    return
##############################

## Define a function to remove duplicates within a list.
def removeDuplicate(list):

    i = 0
    j = 0
    count = 0

    while i < len(list):
        while j < len(list):
            x = list[i]
            y = list[j]
            a = x['term']
            b = y['term']

            if i == j:
                j += 1
            elif a == b:
                list.remove(y)
                count += 1
                j = 0
            else:
                j += 1

# The numbers of duplicates have been counted and stored within the list, if needed.
        x['duplicates'] = count
        i += 1
        j = 0
        count = 0

    return list
##############################

## Load the map data. In JSON format.
## This section will need to be tweaked to increase or decrease the files read.
human = loadJSON()
rat = loadJSON()
fc = loadJSON()

## Extract the required data. Whether it is differences or similarities.
## This section will need to be tweaked depending on what is to be written out to the excel file.
diff1 = comparison('differences', human, fc)
diff2 = comparison('differences', fc, human)
#diff3 = comparison('differences', rat, fc)
#diff4 = comparison('differences', fc, rat)
#diff5 = comparison('differences', human, rat)
#diff6 = comparison('differences', rat, human)

sim1 = comparison('similarities', human, fc)
#sim2 = comparison('similarities', rat, fc)
#sim3 = comparison('similarities', sim1, sim2)

#needFC = removeDuplicate(diff1+diff3)
#needHM = removeDuplicate(diff2+diff6)
#needR = removeDuplicate(diff5+diff4)
###############################

## Prompt the user to select the directory and filename for the excel file.
filename = fd.asksaveasfilename(
        filetypes = (
            ('xlsx files', '*.xlsx'),
            ('csv files', '*.csv'),
            ('All files', '*.*')
        ),
        defaultextension = '.xlsx'
    )

wb = xw.Workbook(filename)

## Write out the data to excel sheets.
## This section will need to be tweaked depending on what is to be written out to the excel file.
writeSheet('Present in all maps', sim1)

writeSheet('Present in AC Human Male', diff1, header='Not in FC')
#addToSheet('Present in AC Human Male', diff5, header='Not in Rat', col=4, row=1)

writeSheet('Present in FC', diff2, header='Not in AC Human Male')
#addToSheet('Present in FC', diff4, header='Not in AC Rat', col=4, row=1)

#writeSheet('Present in AC Rat', diff3, header='Not in FC')
#addToSheet('Present in AC Rat', diff6, header='Not in AC Human Male', col=4, row=1)

writeSheet('Need to add', diff1, header='Need to add to FC')
addToSheet('Need to add', diff2, row=0, col=4, header='Need to add to AC Human Male')
#addToSheet('Need to add', needR, row=0, col=8, header='Need to add to AC Rat')

wb.close()