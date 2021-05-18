# -*- coding: utf-8 -*-
"""
Created on Mon May 10 09:38:10 2021

@author: 4165306
"""
import pandas as pd
import os

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
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


def supprimerColonne(L):
    del L['Nom']
    del L['Prenom']
    del L['DDN']
    del L['S']
    del L['Plaque']
    del L['pos']
    return L


def split(L):
    Lres=L.query('Type == @codeType')
    return Lres


def prep(Lres):
    Lres.insert(4, "RESULTAT", 'prélèvement non conforme', allow_duplicates=False)
    Lres['ID']=Lres['Echantillon']
    return Lres



filename = askopenfilename(title = 'Export Résultat à Pousser')
    
folder = askdirectory(title = 'dossier cible')


os.chdir(folder)

L=extreq(filename)
L = supprimerColonne(L)
Lres=split(L)
Qres=prep(Lres)
if(codeType == "BPO"):
    R=Qres.to_csv('fichier pcr.csv',sep=',')
elif(codeType == "AUT"):
    R=Qres.to_csv('fichier slv.csv',sep=',')