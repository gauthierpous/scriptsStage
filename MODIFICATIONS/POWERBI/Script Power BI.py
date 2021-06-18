# -*- coding: utf-8 -*-
"""
Created on Thu Jun 10 15:07:10 2021

@author: 7032078
"""

import pandas as pd
import os
import numpy as np
import win32com.client as win32

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import Tk
from tkinter.simpledialog import askinteger
from datetime import datetime

Tk().withdraw()

#Demander le type de test
typeTest = 0
codeType = ""
while typeTest != 1 and typeTest != 2:
    typeTest = askinteger(title = 'Quel est le type de test à formater ?', prompt = 'Entrer 1 pour PCR \nEntrer 2 pour Salivaire')

filename = askopenfilename(title = 'Export Résultat à Pousser')
folder = askdirectory(title = 'dossier cible')

Tk().destroy()


#En fonction du chiffre reçu, affecter une valeur à codeType
if(typeTest == 1):
    codeType = "BPO"
elif(typeTest == 2):
    codeType = "AUT"


def extractionDonnees(filename):
    return pd.read_excel(filename)


def split(L):
    if(codeType == "BPO"):
        Lres = L.query('Type == "BPO"')
    if (codeType == "AUT"):
        Lres = L.query('Type == "AUT"')
    return Lres




def formatageDates(dataframe):
    prel = []
    saisie = []
    val = []
    for index, row in dataframe.iterrows():
        prel.append(pd.to_datetime(row['Date Prel'], format='%d/%m/%y %H:%M'))
        saisie.append(pd.to_datetime(row['Date Saisie'], format='%d/%m/%y %H:%M'))
        val.append(pd.to_datetime(row['Date val'], format='%d/%m/%y %H:%M'))
    
    dataframe.loc[:, 'Date Prel'] = prel
    dataframe.loc[:, 'Date Saisie'] = saisie
    dataframe.loc[:, 'Date val'] = val

    return dataframe


def calculDelaiPrescripteur(datesFormates):
    delaisPres = []
    delaisLabo = []
    delaisTotaux = []
    delais = datesFormates.copy()
    for index, row in delais.iterrows():
        delaisPres.append((row['Date Saisie'] - row['Date Prel']) / pd.Timedelta(hours=1))
        delaisLabo.append((row['Date val'] - row['Date Saisie']) / pd.Timedelta(hours=1))
        delaisTotaux.append((row['Date val'] - row['Date Prel']) / pd.Timedelta(hours=1))
        
    delais.insert(9, 'Delai Prescripteur', delaisPres)
    delais.insert(10, 'Delai Laboratoire', delaisLabo)
    delais.insert(11, 'Delai Total', delaisTotaux)
    
    return delais


def numeroSemaine(LresFormate):
    semaines = []
    for i, row in LresFormate.iterrows():
        numero = row["Date Saisie"].isocalendar()[1]
        strNumero = str(numero)
        if numero < 10 :
            strNumero = "0" + strNumero
        semaines.append("Semaine " + strNumero)
    LresFormate.insert(12, 'N° semaine', semaines)
    return LresFormate



#Génère une date pour le nom du fichier
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



#Formate la colonne RESULTATS et la date
def fuscoldoss(Q):
    n=len(Q)
    ind=[i for i in range(n)]
    Q.index=ind
    Q=Q.fillna(0)
    
    C=['Prescripteur',
       'Date', 'Date Prel','Date Saisie','Date val',
       'Type','RESULTAT','Discipline réceptrice', 'MOTIF NC', 
       'Delai Prescripteur', 'Delai Laboratoire', 'Delai Total', 'N° semaine']
    nc = len(C)
    r=np.empty((n,nc),dtype=object)
    for i,row in Q.iterrows():
        r[i,0]=row['Prescripteur']
        r[i,1]=row["Date val"].strftime("%d-%m-%Y")
        r[i,2]=row['Date Prel']
        r[i,3]=row['Date Saisie']
        r[i,4]=row['Date val']
        r[i,5]=row['Type']
        v94=Q.loc[i,'RESULTAT']
        if v94 in ['POSITIF', 'positif','Positif','P']:
            r[i,6]='POS'
        elif v94 in ['*Négatif','Négatif' , 'négatif','N',]:
            r[i,6]='NEG'
        elif v94 in ['I','indeterminé','Indeterminé','Indéterminé']:
            r[i,6]='IND'
        elif v94 in ['prélèvement non conforme','Prélèvement non conforme']:
            r[i,6]='NCONF'
        r[i,7]=row['Discipline réceptrice']  
        r[i,8]=row['MOTIF NC']    
        r[i,9]=row['Delai Prescripteur']
        r[i,10]=row['Delai Laboratoire']
        r[i,11]=row['Delai Total']
        r[i,12]=row['N° semaine']
    return pd.DataFrame(r,columns = C)



#Retourne les statistiques des résultats
def statsResultats(S):
    stats = S.copy()
    nombreLigne = len(stats)
    C=[ 'POSITIF','NEGATIF','INDETERMINE','NON CONFORME',
       'Salicov','Non reçu','Tube fuyant','Volume non respecté','Discordance','Tube vide',
       'Prélèvement d\'expectoration','Contenant non adapté','Absence d\'identité','Autre']
    stats[C] = 0
    for i in range(nombreLigne):
        test = stats["RESULTAT"][i]
        motif = stats["MOTIF NC"][i]
        salicov = stats["Discipline réceptrice"][i]
        if test == 'NCONF':
            stats.loc[i, 'NON CONFORME'] += 1
        elif test == 'POS':   #pos
            stats.loc[i, 'POSITIF'] += 1
        elif test == 'NEG':   #neg
            stats.loc[i, 'NEGATIF'] += 1
        elif test == 'IND':   #ind
            stats.loc[i, 'INDETERMINE'] += 1
        
        if motif == 'prélèvement non reçu':
            stats.loc[i, 'Non reçu'] += 1
        elif motif == 'TUBE FUYANT':
            stats.loc[i, 'Tube fuyant'] += 1
        elif motif == 'volume non respecté':
            stats.loc[i, 'Volume non respecté'] += 1
        elif motif == 'discordance d\'identité':
            stats.loc[i, 'Discordance'] += 1
        elif motif == 'Tube vide':
            stats.loc[i, 'Tube vide'] += 1
        elif motif == 'Prélèvement d\'expectoration':
            stats.loc[i, 'Prélèvement d\'expectoration'] += 1
        elif motif == 'Contenant non adapté':
            stats.loc[i, 'Contenant non adapté'] += 1
        elif motif == 'Absence d\'identité':
            stats.loc[i, 'Absence d\'identité'] += 1
        elif motif == 'Autre':
            stats.loc[i, 'Autre'] += 1

        if "SALICOV" in salicov.upper():
            stats.loc[i, 'Salicov'] += 1
    return stats

  
#Arrondi les dates à 2 décimales pour l'export des données totales
def formatageDate(colonneDonneesTotales):
    for i in range(len(colonneDonneesTotales)):
       colonneDonneesTotales[i] = float(round(colonneDonneesTotales[i], 2))
    return colonneDonneesTotales


#Ajoute une ligne avec le total par colonne 
#Ajoute une colonne aevc le total par ligne
def addtot(R): 
    S=R.copy()
    tot=S.apply(np.sum, axis =0).values
    S.loc['TOTAL']=tot
    
    return S

def supprimeTest(donneesTotales):
    dataFrame = donneesTotales.copy()
    for i, row in dataFrame.iterrows():
        if "TEST" in row["Prescripteur"].upper():
            dataFrame.drop(i, inplace = True)
    return dataFrame


#Ajout des colonnes COVISAN, Autres structures et OPEX
def ajoutStructures(dataframe):
    df = dataframe.copy()
    structures = ['COVISAN', 'Autres structures', 'OPEX']
    covisan = ['COVISAN', 'VALIN', 'DOMICILE']
    autres = ['DIEU', 'ROTHSCHILD', 'CH4V']
    df[structures] = 0
    df["OPEX"] = 1
    for i in range(len(df)):
        if df.loc[i, "Prescripteur"] == 0:
            df.loc[i, "Prescripteur"] = "Prescripteur inconnu"
            
        for elt in covisan:
            if elt in df.loc[i, "Prescripteur"].upper():
                df.loc[i, "COVISAN"] = 1
                df.loc[i, "OPEX"] = 0
        
        for elt1 in autres:
            if elt1 in df.loc[i, 'Prescripteur'].upper():
                df.loc[i, "Autres structures"] = 1
                df.loc[i, "OPEX"] = 0
            
    return df

#Ajout des colonnes 30min et 24h
def ajoutProblemeDelai(dataframe):
    df = dataframe.copy()
    delai = ["30min", "24h"]
    df[delai] = 0
    for i in range(len(df)):
        if df.loc[i, 'Delai Prescripteur'] < 0.5:
            df.loc[i, "30min"] = 1
        if df.loc[i, 'Delai Total'] < 24:
            df.loc[i, "24h"] = 1
            
    return df



#Formate toutes les colonnes de statistiques avec le type int
def page1sepop(S):
    col=['Prescripteur',
         'Date','Type', 
         'Delai Prescripteur','Delai Laboratoire','Delai Total', 'N° semaine',
         'POSITIF' , 'NEGATIF' , 'INDETERMINE' , 'NON CONFORME', 'TOTAL',
         'Salicov','Non reçu','Tube fuyant','Volume non respecté','Discordance','Tube vide',
         'Prélèvement d\'expectoration','Contenant non adapté','Absence d\'identité','Autre',
         'COVISAN', 'Autres structures', 'OPEX',
         "30min", "24h"
         ]
    Q=S[col]
    for elt in col[7:26]:
        Q.loc[:, elt] = Q.loc[:, elt].astype(int) 
    return Q            


#Gère l'exportation du fichier
def expnoind(p1,filename):
    writer=pd.ExcelWriter(filename, engine = 'xlsxwriter')
    #workbook=writer.book
    p1.to_excel(writer, sheet_name='Résultats',index=False)
    writer.save()                    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(folder +"/"+ filename)
    ws1 = wb.Worksheets("Résultats")
    ws1.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


os.chdir(folder)
donnees = extractionDonnees(filename)
dataframeFiltre = split(donnees)

#print("Génère un ID pour chaque résultat")
#identifiant = generationID(dataframeFiltre)

print("Formate les dates")
datesFormates = formatageDates(dataframeFiltre)

print("Calcul des délais des entités et le délais totaux")
delais = calculDelaiPrescripteur(datesFormates)

print("Ajoute le numéro de la semaine de l'acquittement")
Export = numeroSemaine(delais)

print("Restitution d'un dataframe avec l'ajout du champ date")
restitutionFormatee = fuscoldoss(Export)

print("Formatage des délais")
restitutionFormatee['Delai Prescripteur'] = formatageDate(restitutionFormatee['Delai Prescripteur'])
restitutionFormatee['Delai Laboratoire'] = formatageDate(restitutionFormatee['Delai Laboratoire'])
restitutionFormatee['Delai Total'] = formatageDate(restitutionFormatee['Delai Total'])

print('Récupération des statistiques')
donneesTotales = statsResultats(restitutionFormatee)

print("Ajout du total du nombre de tests")
donneesTotales.insert(10, "TOTAL", 1)

print("Ajout des colonnes COVISAN, Autres structures et OPEX")
dataFrame = ajoutStructures(donneesTotales)

print("Ajout des colonnes 30min et 24h")
dataFrame = ajoutProblemeDelai(dataFrame)

print('Formate les colonnes du dataFrame')
dataFrame = page1sepop(dataFrame)

print("Supprime les opérations test")
dataFrame = supprimeTest(dataFrame)

if(codeType == "BPO"):
    dataFrame['Type'] = "PCR"
    expnoind(dataFrame, generer_date()+' Bilan OPS PRC.xlsx')
elif(codeType == "AUT"):
    dataFrame['Type'] = "SLV"
    expnoind(dataFrame, generer_date()+' Bilan OPS SLV.xlsx')