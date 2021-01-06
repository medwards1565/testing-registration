'''
Created on Dec 22, 2020

@author: medwa
'''

# Write a function that accepts an excel file as a parameter (example is in test-registration/resources/) and copies 
# Scheduled Time (column c), First Name (column d), Last Name (column e), Department (column f), Date of Birth (column h) 
# into a new tab.

import openpyxl
from builtins import str

# get our excel file
xlfile = 'C:\\Users\\medwa\\Desktop\\xlfolder\\test.xlsx'

# load workbook
print('test-registration: Loading workbook ' + xlfile)
wb = openpyxl.load_workbook(xlfile)

# create new sheet
tempsheet = 'TEMPSHEET'
print('test-registration: Creating new sheet ' + tempsheet + ' in ' + xlfile)
wb.create_sheet(tempsheet)

# get the sheet to load data from
mastersheet = 'Sheet1'
sheet = wb.get_sheet_by_name(mastersheet)

# copy data
print('test-registration: Copying data from ' + mastersheet + ' to ' + tempsheet + ' in ' + xlfile)
temp_sheet = wb.get_sheet_by_name(tempsheet)
for row in wb [mastersheet] ['B3':'H12']:
    for cell in row:
        # print(cell.coordinate, cell.value)
        temp_sheet[cell.coordinate] = cell.value
    #print('--- END OF ROW ---')

# save the workbook
wb.save('C:\\Users\\medwa\\Desktop\\xlfolder\\test_copy.xlsx')