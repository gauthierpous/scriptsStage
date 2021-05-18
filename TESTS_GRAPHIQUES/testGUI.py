#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  4 12:26:34 2021

@author: gauthierpous
"""

import tkinter as tk
import tkinter.font as font

selection = 0
window = tk.Tk()

def PCR(root):
    global selection
    selection = 1
    verifSelection()
    quitter(root)

def salivaire(root):
    global selection
    selection = 2
    quitter(root)

def TROD(root):
    global selection
    selection = 3
    quitter(root)

def quitter(root):
    root.destroy()

def verifSelection():
    if(selection == 1 or selection == 2 or selection == 3):
        print("La sélection est bien effectuée et elle vaut :", selection)


listOperation = ["PCR", "SALIVAIRE", "TROD"]

window.rowconfigure(0, weight=1, minsize=30)
window.columnconfigure([0, 1, 2], weight=1, minsize=200)
window['bg'] = "#2B2B2B"
window['pady'] = 50

myFont = font.Font(size = 30)

#Premier boutton -- PCR
bouttonPCR = tk.Button(master = window,
                       text = listOperation[0],
                       pady = 20,
                       borderwidth = 2,
                       font = myFont,
                       command = lambda window=window:PCR(window)
                       )
bouttonPCR.grid(row = 0, column = 0, sticky="ew")



#Deuxième boutton -- Salivaire
bouttonSLV = tk.Button(master = window,
                       text = listOperation[1],
                       pady = 20,
                       borderwidth = 2,
                       font = myFont,
                       command = lambda window=window:salivaire(window)
                       )
bouttonSLV.grid(row = 0, column = 1, sticky="ew")



#Troisème boutton -- TROD
bouttonTROD = tk.Button(master = window,
                       text = listOperation[2],
                       pady = 20,
                       borderwidth = 2,
                       font = myFont,
                       command = lambda window=window:TROD(window)
                       )
bouttonTROD.grid(row = 0, column = 2, sticky="ew")

window.mainloop()



