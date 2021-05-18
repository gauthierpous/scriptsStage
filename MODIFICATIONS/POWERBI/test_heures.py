# -*- coding: utf-8 -*-
"""
Created on Tue May 11 15:11:27 2021

@author: 4165306
"""

import pandas as pd
import os

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import Tk
from tkinter.simpledialog import askinteger


Tk().withdraw()

#Demander le type de test
typeTest = 0
codeType = "A"
while typeTest != 1 and typeTest != 2:
    typeTest = askinteger(title = 'Quel est le type de test à formater ?', prompt = 'Entrer 1 pour PCR \nEntrer 2 pour Salivaire')


#En fonction du chiffre reçu, affecter une valeur à codeType
if(typeTest == 1):
    codeType = "BPO"
elif(typeTest == 2):
    codeType = "AUT"


def extreq(filename):
    return pd.read_excel(filename)


def split(L):
    if(codeType == "BPO"):
        Lres = L.query('Type == "BPO"')
    if (codeType == "AUT"):
        Lres = L.query('Type == "AUT"')
    #Lres = Lres.query('RESULTAT == "prélèvement non conforme"')
    return Lres


def prep(Lres):
    Lres['ID']=Lres['Type']+Lres['Prescripteur']+Lres['Date Prel']+Lres['Date Saisie']+Lres['Date val']
    return Lres


def formatageLres(Lres):
    Lres['Date val'] = pd.to_datetime(Lres['Date val'], format='%d/%m/%y %H:%M')
    Lres['Date Saisie'] = pd.to_datetime(Lres['Date Saisie'], format='%d/%m/%y %H:%M')
    Lres['Date Prel'] = pd.to_datetime(Lres['Date Prel'], format='%d/%m/%y %H:%M')
    
    Lres['Delai Prescripteur'] = 0
    Lres['Delai Laboratoire'] = 0
    Lres['Delai Total'] = 0
    return Lres


def calculDelaiPrescripteur(LresFormate):
    delaiPrescripteur = LresFormate.copy()
    for i in delaiPrescripteur.iterrows():
        delaiPrescripteur['Delai Prescripteur'] = (delaiPrescripteur['Date Saisie'] - delaiPrescripteur['Date Prel']) / pd.Timedelta(hours=1)
    return delaiPrescripteur


def calculDelaiLaboratoire(delaiPrescripteur):
    delaiLaboratoire = delaiPrescripteur.copy()
    for i in delaiLaboratoire.iterrows():
        delaiLaboratoire['Delai Laboratoire'] = (delaiLaboratoire['Date val'] - delaiLaboratoire['Date Saisie']) / pd.Timedelta(hours=1)
    return delaiLaboratoire


def calculDelaiTotal(delaiLaboratoire):
    delaiTotal = delaiLaboratoire.copy()
    for i in delaiTotal.iterrows():
        delaiTotal['Delai Total'] = (delaiTotal['Date val'] - delaiTotal['Date Prel']) / pd.Timedelta(hours=1)
    return delaiTotal



filename = askopenfilename(title = 'Export Résultat à Pousser')
folder = askdirectory(title = 'dossier cible')
os.chdir(folder)

L=extreq(filename)
Lres=split(L)
Qres=prep(Lres)
LresFormate = formatageLres(Qres)

delaiPrescripteur = calculDelaiPrescripteur(LresFormate)
delaiLaboratoire = calculDelaiLaboratoire(delaiPrescripteur)
delaiTotal = calculDelaiTotal(delaiLaboratoire)

if(codeType == "BPO"):
    R=delaiTotal.to_csv('Données PCR.csv')
if (codeType == "AUT"):
    R=delaiTotal.to_csv('Données SLV.csv')

