# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 14:44:50 2021

@author: 4178664
"""
import pandas as pd
import os
import numpy as np
import win32com.client as win32
from datetime import datetime
from datetime import date as Date1
import time
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
    
infos=['LOINC' , 'N° Patient BROUSSAIS','Date de naissance', 'Prescripteur', 'Valeur' ]

def generer_date():
    now = datetime.now()
    titre=''
    titre+=str(now.year)
    if now.month<10:
        titre+='0'
    titre+=str(now.month)
    if now.day<10:
        titre+='0'
    titre+=str(now.day)
    return titre

def extractdoss(fichiercsv):
    L=pd.read_csv(fichiercsv,sep=',',header = 1)
    if L.columns[0] != 'Demande':
        L=pd.read_csv(fichiercsv,sep=',')
    return L

def restdoss(L):
    infdoss = ['Prescripteur','Date Prel','Date Saisie','Date Val','Type','Discipline réceptrice','RESULTAT','MOTIF NC']
    Q=L.rename(columns = {'N° patient Ajaccio' :'N° Patient BROUSSAIS' })
        
    Q=Q[['Prescripteur','Date Prel','Date Saisie','Date val','Type','Discipline réceptrice','RESULTAT','MOTIF NC']]

    return Q

def fuscoldoss(Q):
    n=len(Q)
    ind=[i for i in range(n)]
    Q.index=ind
    Q['Res_test']=0
    Q['Symp']=0
    Q=Q.fillna(0)
    C=['Prescripteur','Date Prel','Date Saisie','Date val','Type','Discipline réceptrice','RESULTAT','MOTIF NC']
    nc = len(C)
    r=np.empty((n,nc),dtype=object)
    for i,row in Q.iterrows():
        r[i,0]=row['Prescripteur']
        r[i,1]=row['Date Prel'][0:9]
        r[i,2]=row['Date Saisie']
        r[i,3]=row['Date val']
        r[i,4]=row['Type']
        r[i,5]=row['Discipline réceptrice']  
        r[i,7]=row['MOTIF NC']          
        v94=Q.loc[i,'RESULTAT']
        if v94 in ['POSITIF', 'positif','Positif','P']:
            r[i,6]='POS'
        elif v94 in ['*Négatif','Négatif' , 'négatif','N',]:
            r[i,6]='NEG'
        elif v94 in ['I','indeterminé','Indeterminé','Indéterminé']:
            r[i,6]='IND'
        elif v94 in ['prélèvement non conforme','Prélèvement non conforme']:
            r[i,6]='NCONF'
    return pd.DataFrame(r,columns = C)

    
def presc(S): #renvoie une liste des prescripteurs, sans les tests, et sans doublons
    T=S['Prescripteur'].values
    Presc=[]
    for elt in T:
        if elt not in Presc:
            if 'TEST' not in elt.upper():
                Presc.append(elt)
    return Presc

def datepresc(S):
    Presc = []
    for i , row in S.iterrows():
        dp=row['Date Prel'] + ' - ' + row['Prescripteur']
        if dp not in Presc:
            if 'TEST' not in dp.upper():
                Presc.append(dp)
    return Presc
        
def date(S):
    Date= []
    T=S['Date Prel'].values
    for elt in T:
        if elt not in Date :
            Date.append(elt)
    return Date
        
def statsdate(S,Date):
    C=[ 'POS', 'NEG','IND','NON CONFORME']            
    Index= Date
    Np=len(Date)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    for i,row in S.iterrows():
        ind=row['Date Prel']
        test=row['RESULTAT']
        if ind in Date:
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
    return D

def stats(S,Presc):
    C=[ 'POS', 'NEG','IND','NON CONFORME']            
    Index= Presc
    Np=len(Presc)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    for i,row in S.iterrows():
        ind=row['Prescripteur']
        test=row['RESULTAT']
        if ind in Presc:
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
    return D


def calculate_age(dtob):
    today = Date1.today()
    return today.year - dtob.year - ((today.month, today.day) < (dtob.month, dtob.day))    

def statsdatep(S,Presc):
    C=[ 'POS','NEG','IND','NON CONFORME','Salicov','Non reçu','Tube fuyant','Volume non respecté','Discordance','Tube vide','Prélèvement d\'expectoration','Contenant non adapté','Absence d\'identité','Autre']  
    Index= Presc
    Np=len(Presc)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    for i,row in S.iterrows():
        ind=row['Date Prel'] + ' - ' + row['Prescripteur']
        test=row['RESULTAT']
        motif=row['MOTIF NC']
        salicov= row['Discipline réceptrice']
        if ind in Presc:
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1                    
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
            if motif == 'prélèvement non reçu':
                D.loc[ind]['Non reçu']+=1
            elif motif == 'TUBE FUYANT':
                D.loc[ind]['Tube fuyant']+=1
            elif motif == 'volume non respecté':
                D.loc[ind]['Volume non respecté']+=1
            elif motif == 'discordance d\'identité':
                D.loc[ind]['Discordance']+=1
            elif motif == 'Tube vide':
                D.loc[ind]['Tube vide']+=1
            elif motif == 'Prélèvement d\'expectoration':
                D.loc[ind]['Prélèvement d\'expectoration']+=1
            elif motif == 'Contenant non adapté':
                D.loc[ind]['Contenant non adapté']+=1
            elif motif == 'Absence d\'identité':
                D.loc[ind]['Absence d\'identité']+=1
            elif motif == 'Autre':
                D.loc[ind]['Autre']+=1
            if "SALICOV" in salicov:
                D.loc[ind]['Salicov']+=1
    return D


def addtot(R): #ajoute une ligne avec le total par colonne et par ligne
    S=R.copy()
    tot=S.apply(np.sum, axis =0).values
    S.loc['TOTAL']=tot
    S['TOTAL']=S['POS']+S['NEG']+S['IND']+S['NON CONFORME']
    return S

def page1sepop(S): #stat pos neg pctge 
    col=['Date','Opération','Type', 'POS' , 'NEG' , 'IND' , 'NON CONFORME' , 'TOTAL','Salicov','Non reçu','Tube fuyant','Volume non respecté','Discordance','Tube vide','Prélèvement d\'expectoration','Contenant non adapté','Absence d\'identité','Autre']
    Q=S[col]
    for elt in col[3:17]:
        Q[elt]=Q[elt].astype(int) 
    return Q            

def expnoind(p1,filename):
    writer=pd.ExcelWriter(filename, engine = 'xlsxwriter')
    workbook=writer.book
    p1.to_excel(writer, sheet_name='Résultats',index=False)
    writer.save()                    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(folder +"/"+ filename)
    ws1 = wb.Worksheets("Résultats")
    ws1.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()
    
Tk().withdraw()

fichiercsv=askopenfilename(title = 'sélectionner extraction DOSSIER contenant les ARP')
folder = askdirectory(title = 'sélectionner dossier cible')
Tk().withdraw()
os.chdir(folder)


print("extraction")
L=extractdoss(fichiercsv)
#time.sleep(2)
print("retrait tests")
Q=restdoss(L)
#time.sleep(2)
R=fuscoldoss(Q)
#time.sleep(2)
print("gen liste prescripteurs")
Presc=presc(R)
Datepresc=datepresc(R)
Date = date(R)
print("gen dframe stats/presc")
Stat=stats(R,Presc)
print("gen dframe stats/date+presc")
Statdatep=statsdatep(R,Datepresc)
print("gen dframe stats/date")
Statdate=statsdate(R,Date)
Statdate=Statdate.loc[[elt for elt in Statdate.index if 'TEST' not in elt.upper()]]
#restreint R aux indices covisan
Statdatep=Statdatep.loc[[elt for elt in Statdatep.index if 'TEST' not in elt.upper()]]
#Stat : contient les stats (voir colonnes)

SOPS=Statdatep.loc[[elt for elt in Statdatep.index]]

DFOPS=addtot(SOPS)
INDOP=DFOPS.index.values
DFOPS['Date']=[elt[:9] for elt in INDOP]
DFOPS['Opération']=[elt[12:] for elt in INDOP]
DFOPS['Type']='SLV'

    
DFOPS=page1sepop(DFOPS)
expnoind(DFOPS,generer_date()+' Bilan OPS slv.xlsx')
