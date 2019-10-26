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

print(""" --------- Start of written code ----------
import openpyxl
import os
 ======= End of written code ========== End of written code ======= """)

import openpyxl
import os

print(""" --------- Start of written code ----------
## os.chdir('C:\\Users\\HP\\Desktop') ## For office computer. ##

os.chdir('C:\\Users\\Win-10\\Desktop') ## For home computer. ##
## Always use \\ for avoiding any error in windows.
## use / or // for linux
 ======= End of written code ========== End of written code ======= """)

## os.chdir('C:\\Users\\HP\\Desktop') ## For office computer. ##

os.chdir('C:\\Users\\Win-10\\Desktop') ## For home computer. ##
## Always use \\ for avoiding any error in windows.
## use / or // for linux

print(""" --------- Start of written code ----------
dirpath = os.getcwd()
print("current directory is : " + dirpath)
 ======= End of written code ========== End of written code ======= """)

dirpath = os.getcwd()
print("current directory is : " + dirpath)

print(""" --------- Start of written code ----------

foldername = os.path.basename(dirpath)
print("Directory name is : " + foldername)

scriptpath = os.path.realpath(__file__)
print("Script path is : " + scriptpath)
print ("Current openpyxl version :" , openpyxl.__version__)

 ======= End of written code ========== End of written code ======= """)

foldername = os.path.basename(dirpath)
print("Directory name is : " + foldername)

scriptpath = os.path.realpath(__file__)
print("Script path is : " + scriptpath)

print ("Current openpyxl version :" , openpyxl.__version__)

print(""" --------- Start of written code ----------
wb=openpyxl.load_workbook('C:\\Users\\Win-10\\Desktop\\Testopenpyxl.xlsx')
# type(wb) or print (wb) #
print (wb)
SN=wb.get_sheet_names()
print('sheet_names : '  , SN)
 ======= End of written code ========== End of written code ======= """)

wb=openpyxl.load_workbook('C:\\Users\\Win-10\\Desktop\\Testopenpyxl.xlsx')
# type(wb) or print (wb) #
print (wb)

SN=wb.get_sheet_names()
print('sheet_names : '  , SN)

print(""" --------- Start of written code ----------
sheet= wb.get_sheet_by_name('Sheet1')
# type() or print () #
print (sheet)
 ======= End of written code ========== End of written code ======= """)

sheet= wb.get_sheet_by_name('Sheet1')

# type() or print () #
print (sheet)

print(""" --------- Start of written code ----------
sheet['A1'].value
sheeta1=sheet['A1'].value
print(sheeta1)
sheet['A4'].value
sheeta4=sheet['A4'].value
print(sheeta4)
sheet['A7'].value
sheeta7=sheet['A7'].value
print(sheeta7)

 ======= End of written code ========== End of written code ======= """)

sheet['A1'].value
sheeta1=sheet['A1'].value
print(sheeta1)
sheet['A4'].value
sheeta4=sheet['A4'].value
print(sheeta4)
sheet['A7'].value
sheeta7=sheet['A7'].value
print(sheeta7)

print(""" --------- Start of written code ----------
sheet['B1'].value
sheetB1=sheet['B1'].value
print(sheetB1)
sheet['B4'].value
sheetB4=sheet['B4'].value
print(sheetB4)
sheet['B7'].value
sheetB7=sheet['B7'].value
print(sheetB7)

 ======= End of written code ========== End of written code ======= """)
sheet['B1'].value
sheetB1=sheet['B1'].value
print(sheetB1)
sheet['B4'].value
sheetB4=sheet['B4'].value
print(sheetB4)
sheet['B7'].value
sheetB7=sheet['B7'].value
print(sheetB7)

print(""" ~!@#$%^&* --------- Start of written code ---------- ~!@#$%^&*
sheet['C1'].value
sheetC1=sheet['C1'].value
print(sheetC1)
sheet['C4'].value
sheetC4=sheet['C4'].value
print(sheetC4)
sheet['C7'].value
sheetC7=sheet['C7'].value
print(sheetC7)

~!@#$%^&*  ======= End of written code ========== End of written code =======  ~!@#$%^&* """)
sheet['C1'].value
sheetC1=sheet['C1'].value
print(sheetC1)
sheet['C4'].value
sheetC4=sheet['C4'].value
print(sheetC4)
sheet['C7'].value
sheetC7=sheet['C7'].value
print(sheetC7)


print(""" --------- Start of written code ----------
sheet['C11'].value = 43
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')
## can not save into currently open files.

 ======= End of written code ========== End of written code ======= """)
sheet['C11'].value = 43
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')
## can not save into currently open files.


print(""" --------- Start of written code ----------
STL=sheet.title
print(STL) ## current sheet name

 ======= End of written code ========== End of written code ======= """)
STL=sheet.title
print(STL) ## current sheet name

print(""" --------- Start of written code ----------
sheet.title= "My wb sheet"
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')

 ======= End of written code ========== End of written code ======= """)

sheet.title= "My wb sheet"
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')

print(""" --------- Start of written code ----------
STL2=sheet.title
print(STL2) ## current sheet name

 ======= End of written code ========== End of written code ======= """)
STL2=sheet.title
print(STL2) ## current sheet name

print(""" --------- Start of written code ----------
row_column= sheet.cell(row=1, column=1).value
print(row_column)
 ======= End of written code ========== End of written code ======= """)
row_column= sheet.cell(row=1, column=1).value
print(row_column)

print(""" --------- Start of written code ----------
row_column113= sheet.cell(row=11, column=3).value

## (row=11, column=3) = C11
print(row_column113)

 ======= End of written code ========== End of written code ======= """)

row_column113= sheet.cell(row=11, column=3).value

## (row=11, column=3) = C11
print(row_column113)

print(""" --------- Start of written code ----------
for i in range (1,5):
    print(sheet.cell (row=i, column=3).value)
 ======= End of written code ========== End of written code ======= """)

for i in range (1,5):
    print(sheet.cell (row=i, column=3).value)

print(""" --------- Start of written code ----------
## for knowing maximum numbers of rows and columns
##sheet.max_row
##sheet.max_column
 ======= End of written code ========== End of written code ======= """)
## for knowing maximum numbers of rows and columns
##sheet.max_row
##sheet.max_column


print(""" --------- Start of written code ----------
print(sheet.max_row)
print(sheet.max_column)

print(openpyxl.utils.get_column_letter(1))

 ======= End of written code ========== End of written code ======= """)

print(sheet.max_row)
print(sheet.max_column)

print(openpyxl.utils.get_column_letter(1))



print(openpyxl.utils.get_column_letter(1))


print(""" --------- Start of written code ----------
## print(openpyxl.cell.get_column_letter(1))
## does not work rather use
## print(openpyxl.utils.get_column_letter(1)) ##

 ======= End of written code ========== End of written code ======= """)
## print(openpyxl.cell.get_column_letter(1))
## does not work rather use
## print(openpyxl.utils.get_column_letter(1)) ##

print(""" --------- Start of written code ----------
print(openpyxl.utils.get_column_letter(26))
print(openpyxl.utils.get_column_letter(27))

 ======= End of written code ========== End of written code ======= """)
print(openpyxl.utils.get_column_letter(26))
print(openpyxl.utils.get_column_letter(27))

print(""" --------- Start of written code ----------
print(openpyxl.utils.column_index_from_string('AA'))
print(openpyxl.utils.column_index_from_string('AAA'))
 ======= End of written code ========== End of written code ======= """)

print(openpyxl.utils.column_index_from_string('AA'))
print(openpyxl.utils.column_index_from_string('AAA'))




print("""

--------- Start of written code ---------- --------- Start of written code ----------

wb.create_sheet(title="My 2nd sheet", index=1)
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')

======= End of written code ========== End of written code ======= """)

wb.create_sheet(title="My 2nd sheet", index=1)
wb.save('C:\\Users\\Win-10\\Desktop\\2ndTestopenpyxl.xlsx')

print("""

--------- Start of written code ---------- --------- Start of written code ----------


======= End of written code ========== End of written code =======

  """)



print("""

--------- Start of written code ---------- --------- Start of written code ----------


======= End of written code ========== End of written code =======

  """)



print("""

--------- Start of written code ---------- --------- Start of written code ----------


======= End of written code ========== End of written code =======

  """)



print("""

--------- Start of written code ---------- --------- Start of written code ----------


======= End of written code ========== End of written code =======

  """)


print("""

--------- Start of written code ---------- --------- Start of written code ----------


======= End of written code ========== End of written code =======

  """)





