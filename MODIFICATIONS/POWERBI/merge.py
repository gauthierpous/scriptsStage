# -*- coding: utf-8 -*-
"""
Created on Mon May 10 15:17:38 2021

@author: 4165306
"""

import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.simpledialog import askinteger

Tk().withdraw()



def extreq(filename):
    return pd.read_excel(filename)


#Filtre le dataFrame
def split(L):
    resultatsFiltres=L.query('RESULTAT == "prélèvement non conforme"')
    return resultatsFiltres


#Renome une colonne des resultats
def finaliser(Q):
    Q=Q.rename(columns = {'Discipline réceptrice' : 'Echantillon'})
    return Q[C]


#Merge les deux dataFrame en fonction de leur ID
def stack(nonConformes, resultatsNormes):
    dataframeMerge = nonConformes.merge(resultatsNormes, on='Echantillon', how='right')
    return dataframeMerge


def supprimerColonne(L):
    del L['Nom']
    del L['Prenom']
    del L['DDN']
    del L['S']
    del L['Plaque']
    del L['pos']
    return L


C = ['Pres', 'Date Prel', 'Date Saisie', 'DateVal', 'Echantillon', 'RESULTAT']


ficherNonConformes = askopenfilename(title = 'Fichier des résultats non conformes')
fichierResultats = askopenfilename(title = 'Fichier des résultats')

folder = askdirectory(title = 'Dossier cible')


os.chdir(folder)

nonConformes = extreq(ficherNonConformes)
resultats = extreq(fichierResultats)

resultatsFiltres = split(resultats)

resultatsNormes = finaliser(resultatsFiltres)

dataframeMerge = stack(nonConformes, resultatsNormes)

dataframeFinal = supprimerColonne(dataframeMerge)

dataframeFinal.to_excel('BRS_dossier_compatible_bilanOPS.xlsx')