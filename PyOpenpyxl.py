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

os.chdir('C:\\Users\\HP\\Desktop')

dirpath = os.getcwd()
print("current directory is : " + dirpath)

foldername = os.path.basename(dirpath)
print("Directory name is : " + foldername)

scriptpath = os.path.realpath(__file__)
print("Script path is : " + scriptpath)

print ("Current openpyxl version :" , openpyxl.__version__)





