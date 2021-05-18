# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 14:08:48 2021

@author: 4178664
"""

import pandas as pd
import os

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory


def extreq(filename):
    return pd.read_excel(filename)


def split(L):
    Lres=L.query('Type == "BPO"')
    return Lres

    
def prep(Lres):
    Lres['ID']=Lres['Type']+Lres['Prescripteur']+Lres['Date Prel']+Lres['Date Saisie']+Lres['Date val']
    return Lres



filename = askopenfilename(title = 'Export Résultat à Pousser')
    
folder = askdirectory(title = 'dossier cible')


os.chdir(folder)

L=extreq(filename)
Lres=split(L)
Qres=prep(Lres)
R=Qres.to_csv('fichier pcr.csv',sep=',')