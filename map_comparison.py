import json
from tkinter import *
from tkinter import filedialog as fd
from tkinter import simpledialog as sd
from tkinter import messagebox as mb
import xlsxwriter as xw

## Define a function to load JSON files ##
def loadJSON(mapname):

# Prompt a window to show what needs to be done.
    mb.showinfo(title='Open JSON File', message=('Choose', mapname, 'file'))

    file_path = fd.askopenfilename(
        filetypes= (
            ('JSON files', '*.json'),
            ('All files', '*.*')
        )
    )

    file = open(file_path)
    file_data = json.load(file)

    return file_data
#________________________________________________________________________________________________________________#

## Define a function to save a file in the JSON format ##
def saveFile(data):

# Prompt a window to show what needs to be done.
    mb.showinfo(title='Save JSON File', message=('Choose directory to save JSON file'))

    save_path = fd.asksaveasfilename(
        filetypes = (
            ('JSON files', '*.json'),
            ('All files', '*.*')
        ),
        defaultextension = '.json'
    )
    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
#________________________________________________________________________________________________________________#

## Define a function to sort a dictionary by the label's alphabetical order ##
def sortFn(dict):

    return dict['label']
#________________________________________________________________________________________________________________#

## Define a function to input two lists and extract the differences between the two ##
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
#________________________________________________________________________________________________________________#

## Define a function to input two lists and extract the similarities between the two ##
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
#________________________________________________________________________________________________________________#

## Define a function that inputs two lists, whether the differences or similarities are to be extracted ##
## Note that the 'duplicates' argument is whether to remove duplicates within the list or not ##
def comparison(
        task, list1, list2, duplicates=True
):

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
#________________________________________________________________________________________________________________#

## Define a function to write out to an excel spreadsheet ##
## The required arguments are the worksheet name and data ##
## Optional arguments are a header, the row and column start point, and the titles of the 3 data columns ##
## The 'duplicates' argument is whether the number of occurences of the entry is printed out or not (WIP) ##
def writeSheet(
        worksheet, data, header=None, row=0, col=0, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=False
):

    ws = wb.add_worksheet(worksheet)

# Add bold format
    bold = wb.add_format({'bold': True})

# Add format to change font colour to bold red and size column width.
    red = wb.add_format({'bold': True, 'font_color': 'red'})

# The header will be printed in the desired start column.
    if header != None:
        ws.write(0, col, header.upper(), red)
        if len(header) > 17:
           ws.set_column(0, col, 30)
        else:
            ws.set_column(0, col, 22)
        col += 1

# Write out the data titles and size column widths.
    ws.write(0, col, A1Title, bold)
    ws.set_column(col, col, 35)
    ws.write(0, col + 1, B1Title, bold)
    ws.set_column((col + 1), (col + 1), 15)
    ws.write(0, col + 2, C1Title, bold)
    ws.set_column((col + 2), (col + 2), 15)

# Write out the title for the count of duplicates, if included.
    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates', bold)

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
#________________________________________________________________________________________________________________#

## Define a function to write out to an existing excel spreadsheet ##
## The required arguments are the worksheet name, data, and starting row and column ##
## The optional arguments are the header name and data titles ##
## The duplicates argument is to whether to include the count of duplicates in the data or not ##
def addToSheet(
        worksheet, data, row, col, header=None, A1Title='Label', B1Title='Term', C1Title='ID', duplicates=False
):

    ws = wb.get_worksheet_by_name(worksheet)

# Add bold format for headers and titles
    bold = wb.add_format({'bold': True})

# Add format to change font colour to red.
    red = wb.add_format({'bold': True, 'font_color': 'red'})

# Write out the header in the starting column.
    if header != None:
        ws.write(0, col, header.upper(), red)
        if len(header) > 16:
           ws.set_column(0, col, 30)
        else:
            ws.set_column(0, col, 22)

        col += 1

# Write out the data titles and size columns.
    ws.write(0, col, A1Title, bold)
    ws.set_column(col, col, 35)
    ws.write(0, col + 1, B1Title, bold)
    ws.set_column((col + 1), (col + 1), 15)
    ws.write(0, col + 2, C1Title, bold)
    ws.set_column((col + 2), (col + 2), 15)

# Ensure the data titles in row 0 are not overwritten.
    if row == 0:
        row += 1

# Write out the title for the count of duplicates, if included.
    if duplicates == True:
        ws.write(0, col + 3, 'Duplicates', bold)
        ws.set_column((col+3), (col+3), 10)

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
#________________________________________________________________________________________________________________#

## Define a function to remove duplicates within a list ##
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
#________________________________________________________________________________________________________________#
def ordinal(n: int):
    if 11 <= (n % 100) <= 13:
        suffix = 'th'
    else:
        suffix = ['th', 'st', 'nd', 'rd', 'th'][min(n % 10, 4)]

        return str(n) + suffix
#________________________________________________________________________________________________________________#
def desire(input):
    global option
    option = input
    win.quit()
    win.destroy()
    win.update()

    return option

#________________________________________________________________________________________________________________#
def comparisonOption():
    win = Toplevel()
    win.title('Desired comparison')
    Label(win, text='What comparison would you like?').pack()
    Button(win, text='Differences').pack()


#________________________________________________________________________________________________________________#

## Load maps.

# Prompt user to input desired number of maps.
quantity = int(sd.askstring('Number of maps', 'How many maps are to be loaded?'))

i = 1
m = {}

# Loop through desired number of maps to input name
while i <= quantity:
    n = ordinal(i)
    question = 'What is the name of the', n, 'map?'
    map = sd.askstring('Map name', question)
    map.replace(" ", "_")

# Use the loadJSON function to load each map.
    for x in range(i , (i + 1)):
        m["map{0}".format(x)] = loadJSON(map)
        break

    i += 1

#________________________________________________________________________________________________________________#

## Extract the required data. Whether it is differences or similarities.

# Prompt user input for what is desired - differences, similarities, or both.
root = Tk()
win = Toplevel()
win.geometry('200x150')
win.title('Desired comparison')
Label(win, text='What comparison would you like?').pack()
Button(win, text='Differences', command=lambda *args: desire('differences')).pack()
Button(win, text='Similarities', command=lambda *args: desire('similarities')).pack()
Button(win, text='Both', command=lambda *args: desire('both')).pack()

root.withdraw()
win.mainloop()

# Execute the desired process.
i = 1
z = 0
d = {}
s = {}

while i <= quantity:
    for p in range(i, (i+1)):
        list1 = list(m.keys())[z]
        list2 = list(m.keys())[z + 1]
        if option == 'differences':
            d["diff{0}".format(p)] = comparison('differences', m[list1], m[list2])
            d["diff{0}".format(p+1)] = comparison('differences', m[list2], m[list1])
        elif option == 'similarities':
            s["sim{0}".format(p)] = comparison('similarities', m[list1], m[list2])
        elif option == 'both':
            d["diff{0}".format(p)] = comparison('differences', m[list1], m[list2])
            d["diff{0}".format(p + 1)] = comparison('differences', m[list2], m[list1])
            s["sim{0}".format(p)] = comparison('similarities', m[list1], m[list2])

    i += 1

#________________________________________________________________________________________________________________#

## Prompt the user to select the directory and filename for the excel file ##

# Prompt a window to show what needs to be done.
mb.showinfo(title='Save excel File', message=('Choose directory and name to save excel file'))

filename = fd.asksaveasfilename(
        filetypes = (
            ('xlsx files', '*.xlsx'),
            ('csv files', '*.csv'),
            ('All files', '*.*')
        ),
        defaultextension = '.xlsx'
    )

wb = xw.Workbook(filename)

## Write out the data to excel sheets ##
## This section will need to be tweaked depending on what is to be written out to the excel file ##
writeSheet('Present in all maps', sim1)

writeSheet('Present in AC Human Male', diff1, header='Not in FC')
writeSheet('Present in FC', diff2, header='Not in AC Human Male')
writeSheet('Need to add', diff1, header='Need to add to FC')
addToSheet('Need to add', diff2, row=0, col=4, header='Need to add to AC Human Male')

wb.close()

#________________________________________________________________________________________________________________#