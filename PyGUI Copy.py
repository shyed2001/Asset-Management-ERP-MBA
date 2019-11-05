#-------------------------------------------------------------------------------
# Name:        PyGUI Copy PYGUIautoloader
# Purpose:     PYGUIautoloader Automate APP programs loading
#
# Author:      Shyed Shahriar Housaini
#
# Created:     05/11/2019
# Copyright:   (c) Shyed Shahriar Housaini 2019
# Licence:     <your licence: Shyed Shahriar Housaini's terms and conditions.>
#-------------------------------------------------------------------------------

## PYGUIautoloader Automate APP programs loading ##

import tkinter as tk
from tkinter import filedialog, Text 
import os

root = tk.Tk()
canvas = tk.Canvas(root, height=700, width=700, bg="#263D42")
canvas.pack()

frame = tk.Frame(root, bg="white")
frame.place(relwidth=0.8, relheight=0.8)

root.mainloop()



