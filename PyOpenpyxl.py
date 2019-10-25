#-------------------------------------------------------------------------------
# Name:        PyOpenpyxl.py
# Purpose:
#
# Author:      Shyed Shahriar Housaini
#
# Created:     24/10/2019
# Copyright:   (c) Shyed Shahriar Housaini 2019
# Licence:     <your licence: Shyed Shahriar Housaini's terms and conditions.>
#-------------------------------------------------------------------------------

import openpyxl
import os

## os.chdir('C:\\Users\\HP\\Desktop') ## For office computer. ##

os.chdir('C:\\Users\\Win-10\\Desktop') ## For home computer. ##

dirpath = os.getcwd()
print("current directory is : " + dirpath)

foldername = os.path.basename(dirpath)
print("Directory name is : " + foldername)

scriptpath = os.path.realpath(__file__)
print("Script path is : " + scriptpath)

print ("Current openpyxl version :" , openpyxl.__version__)

wb=openpyxl.load_workbook('C:\\Users\\Win-10\\Desktop\\Testopenpyxl.xlsx')
# type(wb) or print (wb) #
#print (wb)

SN=wb.get_sheet_names()
print('sheet_names : '  , SN)

sheet= wb.get_sheet_by_name('Sheet1')

# type() or print () #
print (sheet)


sheeta1=sheet['A1'].value

print(sheeta1)







