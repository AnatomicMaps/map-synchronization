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


def fileTypes():

    filetypes = (
        ('JSON files', '*.json'),
        ('All files', '*.*')
    )

    return filetypes
##############################

def loadJSON():

    file_path = fd.askopenfilename(
        filetypes=fileTypes()
    )

    file = open(file_path)
    file_data = json.load(file)

    return file_data
##############################

def saveFile(data):

    save_path = fd.asksaveasfilename(
        filetypes = fileTypes(),
        defaultextension = '.json'
    )
    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
##############################

def sortFn(dict):

    return dict['label']
##############################

def getDifference(file1, file2):
    i = 0
    j = 0
    diff = 0
    differences = []

    while i < len(file1):
        while j < len(file2):
            if i == len(file1):
                break
            else:
                x = file1[i]
                y = file2[j]

                if x['term'] == y['term']:
                    i += 1
                    diff = 0
                    j = 0
                else:
                    diff += 1
                    j += 1

        if diff > 0:
            differences.append(file1[i])

        i += 1
        diff = 0
        j = 0

    differences.sort(key=sortFn)

    return differences
##############################

def getSimilar(file1, file2):
    i = 0
    j = 0
    similarities = []

    while i < len(file1):
        while j < len(file2):
            x = file1[i]
            y = file2[j]

            if x['term'] == y['term']:
                similarities.append(file1[i])
                j += 1
            else:
                j += 1

        i += 1
        j = 0

    similarities.sort(key=sortFn)

    return similarities
##############################

def comparison(task, file1, file2, duplicates = True):
    list = []
    new_list = []

    if task == 'differences':
        list = getDifference(file1, file2)
    elif task == 'similarities':
        list = getSimilar(file1, file2)
    else:
        print('Error')

    if duplicates == True:
        new_list = removeDuplicate(list)
        return new_list
    else:
        return list
##############################

def writeSheet(worksheet, data, header=None, col=0, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=True):

    ws = wb.add_worksheet(worksheet)

    if header != None:
        ws.write(0, 0, header)
        if col == 0:
            col += 1

    ws.write(0, col, A1Title)
    ws.write(0, col + 1, B1Title)
    ws.write(0, col + 2, C1Title)

    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates')

    i = 0
    row = 1

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

def addToSheet(worksheet, data, header=None, col=5, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=True):

    ws = wb.get_worksheet_by_name(worksheet)

    if header != None:
        ws.write(0, col, header)
        col += 1

    ws.write(0, col, A1Title)
    ws.write(0, col + 1, B1Title)
    ws.write(0, col + 2, C1Title)

    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates')

    i = 0
    row = 1

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

        x['duplicates'] = count
        i += 1
        j = 0
        count = 0

    return list
##############################

human = loadJSON()
rat = loadJSON()
fc = loadJSON()

diff1 = comparison('differences', human, fc)
diff2 = comparison('differences', fc, human)
diff3 = comparison('differences', rat, fc)
diff4 = comparison('differences', fc, rat)
diff5 = comparison('differences', human, rat)
diff6 = comparison('differences', rat, human)

sim1 = comparison('similarities', human, fc)
sim2 = comparison('similarities', rat, fc)
sim3 = comparison('similarities', sim1, sim2)

###############################

filename = fd.asksaveasfilename(
        filetypes = (
            ('xlsx files', '*.xlsx'),
            ('All files', '*.*')
        ),
        defaultextension = '.xlsx'
    )

wb = xw.Workbook(filename)

writeSheet('Present in all maps', sim3, duplicates=False)

writeSheet('Present in AC Human Male', diff1, header='Not in FC', duplicates=False)
addToSheet('Present in AC Human Male', diff5, header='Not in Rat', col=4, duplicates=False)

writeSheet('Present in FC', diff2, header='Not in AC Human Male', duplicates=False)
addToSheet('Present in FC', diff4, header='Not in AC Rat', col=4)

writeSheet('Present in AC Rat', diff3, header='Not in FC', duplicates=False)
addToSheet('Present in AC Rat', diff6, header='Not in AC Human Male', duplicates=False, col=4)

writeSheet('Need to add', diff1, header='Need to add to FC', duplicates=False)
addToSheet('Need to add', diff3, row=len(diff1), duplicates=False, col=4)

addToSheet('Need to add', diff2, header='Need to add to AC Human Male', col=4, duplicates=False)
addToSheet('Need to add', diff6, row=len(diff2), col=4, duplicates=False)

addToSheet('Need to add', diff5, header='Need to add to rat', duplicates=False, col=8)
addToSheet('Need to add', diff4, row=len(diff5), duplicates=False, col=8)

wb.close()