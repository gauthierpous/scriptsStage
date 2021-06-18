# -*- coding: utf-8 -*-
"""
Created on Mon Jun 14 11:35:44 2021

@author: 7032078
"""

import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory


#Extraction des données
def extreq(filename):
    return pd.read_excel(filename)


#Génération d'un ID (à changer si on peut faire l'export du numéro d'échantillon)
def generationID(df):
    identifiants = []
    for i, row in df.iterrows():
        typetest = str(row['Type'])
        prel = str(row['Date Prel'])
        saisie = str(row['Date Saisie'])
        nom = str(row['Nom'])
        identifiants.append(typetest + prel + saisie + nom)
    df.insert(8, 'identifiant', identifiants)
    return df


#Fusionne les deux dataFrame en fonction de leur ID
def stack(nonConformes, resultats):
    dataframeMerge = nonConformes.merge(resultats, on='identifiant', how='right')
    return dataframeMerge


#Renome les colonnes
def rename(df):
    df = df.rename(columns = {'Presc_y' : 'Prescripteur'})
    df = df.rename(columns = {'Date Prel_y' : 'Date Prel'})
    df = df.rename(columns = {'Date Saisie_y' : 'Date Saisie'})
    df = df.rename(columns = {'Date val_y' : 'Date val'})
    df = df.rename(columns = {'Type_y' : 'Type'})
    return df



Tk().withdraw()

ficherNonConformes = askopenfilename(title = 'Fichier des résultats non conformes')
fichierResultats = askopenfilename(title = 'Fichier des résultats')
folder = askdirectory(title = 'Dossier cible')

Tk().destroy()


os.chdir(folder)

print("Extraction des résultats dans une dataframe")
dfNonConformes = extreq(ficherNonConformes)
dfResultats = extreq(fichierResultats)

print("Génération des identifiants dans les deux tableaux")
resultats = generationID(dfResultats)
nonConformes = generationID(dfNonConformes)

print("Réalisation de la fusion")
fusion = stack(nonConformes, resultats)

print("Mise à jour des noms des colonnes")
dataFrame = rename(fusion)

print("Génération du dataframe final avec les bonnes colonnes")
export = dataFrame[["Prescripteur", "Date Prel", "Date Saisie", "Date val", "Type", "Discipline réceptrice", "RESULTAT", "MOTIF NC"]]


export.to_excel('Fusion_Résultats_Non_Conformes.xlsx')