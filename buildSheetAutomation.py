from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Color, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
from subprocess import check_output
from tkcolorpicker import askcolor
import tkinter as tk
import tkinter.ttk as ttk
import string
import re
import os
home = os.path.expanduser('~')
doc = home + '/Documents/testBuildSheet.xlsx'
wb = Workbook()
sheet = wb.active

lsideBorders = Border(left=Side(style='thick'))
rsideBorders = Border(right=Side(style='thick'))
bBorders = Border(bottom=Side(style='thick'))
topBorders = Border(top=Side(style='thick'))
tlBorder = Border(top=Side(style='thick'), left=Side(style='thick'))
trBorder = Border(top=Side(style='thick'), right=Side(style='thick'))
blBorder = Border(bottom=Side(style='thick'), left=Side(style='thick'))
brBorder = Border(bottom=Side(style='thick'), right=Side(style='thick'))
iipBorderl = Border(top=Side(style='thin'), bottom=Side(style='thin'), left=Side(style='thin'))
iipBorderlt = Border(top=Side(style='thick'), bottom=Side(style='thin'), left=Side(style='thin'))
iipBorderrt = Border(top=Side(style='thick'), bottom=Side(style='thin'), right=Side(style='thin'))
iipBorderr = Border(top=Side(style='thin'), bottom=Side(style='thin'), right=Side(style='thin'))
backgroundColor = PatternFill(start_color='ffcc99', end_color='ffcc99', fill_type='solid')
titleFill = PatternFill(start_color='ccffff', end_color='ccffff', fill_type='solid')
titleFont = Font(name='Calibri', size=14, bold=True)

cwidths = {'A': 2.57, 'B': 23, 'C': 20, 'D': 15.57, 'E': 17.71, 'F': 16.29, 'G': 13, 'H': 10.57, 'I': 12.43, 'J': 10, 'K': 19.71, 'L': 8.57}
headings = ['System', 'Type', 'Name', 'Interface', 'IP', 'User ID', 'PWD', 'Console', 'State', 'Switch::port#', 'Rack']
yesList = ['YES', 'Y']
noList = ['NO', 'N']
dupList = ['DUP', 'DUPLICATE', 'COPY', 'SAME']
portList = list()
newLocation = 5
name = ''
backupName=''
loc = ''
#instruction for duplicate
dup = list()


def setBackgroundColor(c):
    try:
        global backgroundColor
        backgroundColor = PatternFill(start_color=c, end_color=c, fill_type='solid')
    except:
        print("Wrong color")


"""
addDeviceLayout()
Arguments:
    rs: Row start. At what row the device layout begins
    rend: Row end. At what row th device layout ends
Description:
Formats the cells in excel to create the device entry.

"""


def addDeviceLayout(rs, rend):
    t = 0
    for y in range(2, 13):
        for x in range(rs, rend+1):
            sheet.cell(row=x, column=y).fill = backgroundColor
            #first row
            if x == rs:
                # if y == 4:
                #     sheet.cell(row=x, column=y).value=name
                if y == 2:
                    sheet.cell(row=x, column=y).value=name
                    sheet.cell(row=x, column=y).border = tlBorder
                elif y == 11 or y == 12:
                    sheet.cell(row=x, column=y).border = trBorder
                else:
                    sheet.cell(row=x, column=y).border = topBorders
            #last row
            elif x == rend:
                if y == 2:
                    sheet.cell(row=x, column=y).border = blBorder
                elif y == 11 or y == 12:
                    sheet.cell(row=x, column=y).border = brBorder
                else:
                    sheet.cell(row=x, column=y).border = bBorders
            #all other rows
            else:
                if y == 2:
                    sheet.cell(row=x, column=y).border = lsideBorders
                elif y == 11 or y == 12:
                        sheet.cell(row=x, column=y).border = rsideBorders
        #area for ports
        for x in range(rs, rend):  
            if y == 5:
                if x == rs:
                    sheet.cell(row=x, column=y).border = iipBorderlt
                else:
                    sheet.cell(row=x, column=y).border = iipBorderl
                if len(portList) >= t:
                    sheet.cell(row=x, column=y).value = portList[t]
                t += 1
            elif y == 6:
                if x == rs:
                    sheet.cell(row=x, column=y).border = iipBorderrt
                else:
                    sheet.cell(row=x, column=y).border = iipBorderr
    sheet.merge_cells(start_row=rs, start_column=12, end_row=rend, end_column=12)
    sheet.cell(row=rs, column=12).alignment= Alignment(horizontal='center', vertical='center')
    sheet.cell(row=rs, column=12).value = loc    


"""
makeHeader()
Arguments:
    cs: Column Start (usually 4 for the header)
    cend: Column End (usually 13 for the header)
    r: Row number that the header will be placed on
Description:
Formats the cells in excel to create a header

"""


def makeHeader(r, cs=2, cend=12):
    t = 0
    for y in range(cs, cend+1):
        sheet.cell(row=r, column=y).fill = titleFill
        sheet.cell(row=r, column=y).border = topBorders
        if y == cs:
            sheet.cell(row=r, column=y).border = tlBorder
        elif y == cend-1 or y == cend:
            sheet.cell(row=r, column=y).border = trBorder
        if t < len(headings):
            sheet.cell(row=r, column=y).value = headings[t]
        t += 1
        sheet.cell(row=r, column=y).font = titleFont

"""
chooseColo()
Arguments: none
Description:
    Uses tkinter color selection form to allow user to select color they want the device information
    to appear in.
"""


def chooseColor():
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use('clam')
    root.lift()
    root.attributes('-topmost', True)
    root.focus_force()
    
    color = askcolor((243,205,121), root)[1]
    if color != '' and color != None:
        color = color[1:]
        global backgroundColor
        backgroundColor = PatternFill(start_color=color, end_color=color, fill_type='solid')
    root.destroy()
    root.mainloop()


"""
getPorts()
Arguments:None
Description:
    Asks user for list of ports and parses returned information.
"""


def getPorts():
    global portList
    ports = input("Ports on device (ex. e0a, e0b | e0a-e0z | Port 1-Port 10): ")
    if ',' in ports:
        p = ports.split(',')
    else:
        p = {ports}

    for i in p:
            if '-' not in i:
                portList.append(i)
            
            elif i[0] == 'e' or i[1] == 'e':
                beginningList = list()
                endList = list()
                b=0
                e=0
                x = re.search(r"[e][0-9][a-z][-][e][0-9][a-z]", i)
                if x != None:
                    entry = x.string.split('-')
                    beginningList.append(entry[0][-3] + entry[0][-2])
                    beginningList.append(entry[1][-3] + entry[1][-2])
                    endList.append(entry[0][-1])
                    endList.append(entry[1][-1])
                    if beginningList[0] == beginningList[1]:
                        start = 0
                        stop = 0
                        for l in string.ascii_lowercase:
                            if l == endList[0]:
                                start = string.ascii_lowercase.index(l)
                            elif l == endList[1]:
                                stop = string.ascii_lowercase.index(l)
                        if start == stop:
                            portList.append(beginningList[0] + endList[0])
                        elif start > stop:
                            e = start
                            b = stop
                        elif start < stop:
                            b = start
                            e = stop
                        for x in range(b+1, e):
                            endList.append(string.ascii_lowercase[x])
                        endList.sort()
                        for e in endList:
                            portList.append(beginningList[0] + e)
            elif i[-1] in string.digits:
                x = re.search(r"[a-zA-Z][ ]*[0-9]+", i)
                if x != None:
                    letterList = list()
                    endpoints = list()
                    counter = 0
                    beginning = 0
                    end = 0
                    findRange = x.string.split('-')
                    #gets first letter from both sides of '-'
                    for r in findRange:
                            letterList.append(str(''.join(list(filter(str.isalpha, r)))))
                    #checks if there is a letter to the right of '-', and if not simply takes the number
                    if letterList[1] is not None and letterList[1] != "":
                            if letterList[0].upper() == letterList[1].upper():
                                    for r in findRange:
                                            endpoints.append(int(''.join(list(filter(str.isdigit, r)))))
                    elif letterList[1] is None or letterList[1] == "":
                            for r in findRange:
                                    endpoints.append(int(''.join(list(filter(str.isdigit, r)))))
                    #compares the 2 numbers so that they can be traversed even if in the wrong order
                    if endpoints != None:
                            if endpoints[1] > endpoints[0]:
                                    end = endpoints[1]
                                    beginning = endpoints[0]
                            elif endpoints[0] > endpoints[1]:
                                    beginning = endpoints[1]
                                    end = endpoints[0]
                            counter = end - beginning
                    #adds the first entry to the list [letter][number]
                    if counter == 0 and (letterList is not None or endpoints is not None):
                            portList.append(letterList[0] + endpoints[0])
                    #cycles through to add the rest of the entries
                    elif counter > 0:
                            for x in range(beginning, end + 1):
                                y = re.search(r"[a-zA-Z][ ][0-9]+", i)
                                if y != None:
                                    portList.append(letterList[0] + " {}".format(x))
                                else:
                                    portList.append(letterList[0] + "{}".format(x))



fname = input("What would you like the filename to be? ")
doc= home + '/Documents/{}.xlsx'.format(fname)


for c, s in cwidths.items():
    sheet.column_dimensions[c].width = s

makeHeader(4)

d = False
l = True
runtimes = 0
while l == True:
    if d == False:
        name = input("Type of device: ").upper()
    if runtimes > 0:
        r = True
        if d == False:
            while r == True:
                keepPorts = input("Same ports as last time? ")
                if keepPorts.upper() in noList:
                    portList.clear()
                    getPorts()
                    r = False
                elif keepPorts.upper() in yesList:
                    print("Keeping same ports: " + str(portList))
                    r = False
                else:
                    print("Incorrect response(Y/N needed)\n")
        else:
            print("Keeping same ports: " + str(portList))
    else:
        getPorts()
    if d == False:
        loc = input("Device rack location: ")
        cloop = True
        while cloop == True:
            col = input("Choose color for device? ")
            if col.upper() in yesList:
                chooseColor()
                cloop = False
            if col.upper() in noList:
                cloop = False
    addDeviceLayout(newLocation, len(portList) + newLocation)
    newLocation += (len(portList) + 1)
    loop = True
    while loop == True:
        begin = input("Add a device? ")
        if begin.upper() in noList:
            l = False
            loop = False
        elif begin.upper() in yesList:
            d = False
            runtimes += 1
            loop = False
        elif begin.upper() in dupList:
            d = True
            runtimes += 1
            loop = False
        else:
            print("Incorrect response(Y/N/Dup needed)\n")




wb.save(doc)
check_output(doc, shell=True)