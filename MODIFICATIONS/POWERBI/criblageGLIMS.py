# -*- coding: utf-8 -*-
"""
Created on Tue Jun 15 09:47:53 2021

@author: 7032078
"""

import pandas as pd
import os

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory



def extreq(filename):
    return pd.read_excel(filename)


#Restitue le fichier CSV sous une dataFrame
def restitution(L):    
    colonne = ['Analyse','Prescripteur',
       'Date de prélèvement de dossier', 'Heure de prélèvement de dossier',
       'Valeur','Description du groupe']
    Q = L[colonne]
    return Q


def dropDonnees(df):
    indexNames = df[ df['Rens 2'] == 'VAR_SUSP' ].index
    df.drop(indexNames , inplace=True)
    return df


def formatageDates(df):
    dates = []
    for i, row in df.iterrows():
        dates.append(row["Date val"][0:8]) 
    df.insert(1, "Date", dates)
    return df

Tk().withdraw()

filename = askopenfilename(title = 'Export Cyberlab')    
folder = askdirectory(title = 'Dossier cible')

os.chdir(folder)

Tk().destroy()


print("Extraction des données depuis l'export")
donneesBrutes = extreq(filename)

#print("Restitution sous forme de dataFrame")
#dataframe = restitution(donneesBrutes)

print("Suppression des lignes avec VARSUSP")
drop = dropDonnees(donneesBrutes)

print("Formatage des dates")
date = formatageDates(drop)


Export = date.to_excel('Criblage.xlsx')
