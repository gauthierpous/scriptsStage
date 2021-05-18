# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 14:08:48 2021

@author: 4178664
"""

import pandas as pd
import os
import numpy as np
import io
from datetime import time

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.simpledialog import askstring


def extreq(filename):
    return pd.read_excel(filename)

def split(L):
    Lres=L.query('Type == "AUT"')
    
    return Lres
    
def prep(Lres):
    Lres['ID']=Lres['Type']+Lres['Prescripteur']
    return Lres

def stack(Qres,Qsym,Qheb):
    Q=Qres.merge(Qsym,left_on='ID',right_on='ID')
    Q=Q.merge(Qheb,left_on='ID',right_on='ID')
    return Q

def finaliser(Q):
    Q=Q.rename(columns = {'Valeur' : 'TYPOR'})
    Q=Q.rename(columns = {'Valeur_x' : '94845-5'})
    Q=Q.rename(columns = {'Valeur_y' : 'APSYM'})
    Q=Q.rename(columns = {"Date du dernier compte-rendu de résultat" : 'Date_CR'})
    return Q[C]

def export(Q):
    Ldate=list(set(Q['Date_CR'].values))
    for elt in Ldate:
        Qd=Q.query('Date_CR == @elt')
        elt=elt.replace('/','-')
        Qd.to_csv(elt+'_BRS.csv',index = False,sep=';')

C=['Nom','Presc']


filename = askopenfilename(title = 'Export Résultat à Pousser')
    
folder = askdirectory(title = 'dossier cible')


os.chdir(folder)

L=extreq(filename)
#Lres,Lsym,Lheb=split(L)
Lres=split(L)
#Qres,Qsym,Qheb=prep(Lres,Lsym,Lheb)
Qres=prep(Lres)
#Q=stack(Qres,Qsym,Qheb)
R=Qres.to_csv('fichier slv.csv',sep=',')



