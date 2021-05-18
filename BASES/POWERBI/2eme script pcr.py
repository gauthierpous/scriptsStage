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
from statistics import mean
#import time
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory

#Je pense que ça ne sert à rien    
infos=['LOINC' , 'N° Patient BROUSSAIS','Date de naissance', 'Prescripteur', 'Valeur' ]

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

#Récupère le fichier CSV
def extractdoss(fichiercsv):
    L=pd.read_csv(fichiercsv,sep=',',header = 1)
    if L.columns[0] != 'Demande':
        L=pd.read_csv(fichiercsv,sep=',')
    return L

#Restitue le fichier CSV sous un autre nom de variable... Pas très utile
def restdoss(L):
    #infdoss = ['Prescripteur','Date Prel','Date Saisie','Date Val','Type','Discipline réceptrice','RESULTAT','MOTIF NC']
    Q=L.rename(columns = {'N° patient Ajaccio' :'N° Patient BROUSSAIS' })
   
    #Q=L parce que la ligne d'au-dessus ne sert à rien   
    Q[['Prescripteur','Date Prel','Date Saisie','Date val','Type','Discipline réceptrice','RESULTAT','MOTIF NC',
       'Delai Prescripteur', 'Delai Laboratoire', 'Delai Total']]
    return Q


def fuscoldoss(Q):
    n=len(Q)
    ind=[i for i in range(n)]
    Q.index=ind
    Q['Res_test']=0
    Q['Symp']=0
    Q=Q.fillna(0)
    
    C=['Prescripteur','Date Prel','Date Saisie','Date val','Type','Discipline réceptrice','RESULTAT','MOTIF NC', 
       'Delai Prescripteur', 'Delai Laboratoire', 'Delai Total']
    nc = len(C)
    r=np.empty((n,nc),dtype=object)
    for i,row in Q.iterrows():
        r[i,0]=row['Prescripteur']
        r[i,1]=row['Date Prel'][0:10]
        r[i,2]=row['Date Saisie']
        r[i,3]=row['Date val']
        r[i,4]=row['Type']
        r[i,5]=row['Discipline réceptrice']  
        r[i,7]=row['MOTIF NC']    
        r[i,8]=row['Delai Prescripteur']
        r[i,9]=row['Delai Laboratoire']
        r[i,10]=row['Delai Total']
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

#Renvoie une liste des prescripteurs, sans les tests, et sans doublons
def presc(S): 
    T=S['Prescripteur'].values
    Presc=[]
    for elt in T:
        if elt not in Presc:
            if 'TEST' not in elt.upper():
                Presc.append(elt)
    return Presc

#Renvoie une liste des dates de prescription, sans les tests, et sans doublons
def datepresc(S):
    Presc = []
    for i , row in S.iterrows():
        dp=row['Date Prel'] + ' - ' + row['Prescripteur']
        if dp not in Presc:
            if 'TEST' not in dp.upper():
                Presc.append(dp)
    return Presc


#Renvoie la liste des moyennes des prescripteurs et des laboraoire en fonction des date de prélèvement
#Arrondi les durées à 2 décimales
def moyennes(S):
    moyennesPrescripteurs = []
    moyennesLaboratoire = []
    moyennesTotal = []
    prescripteurActuel = ''
    heurePrescripteurTotal = 0
    heureLaboratoireTotal = 0
    heureTotal = 0
    compteur = 0
    
    for i , row in S.iterrows():
        prescripteur = row['Date Prel'] + ' - ' + row['Prescripteur']
        
        if(i == 0):
            prescripteurActuel = prescripteur
            
        if(prescripteur != prescripteurActuel):
            prescripteurActuel = row['Date Prel'] + ' - ' + row['Prescripteur']
            moyenneHeuresPrescripteur = float(round(heurePrescripteurTotal / compteur, 2))
            moyenneHeuresLaboratoire = float(round(heureLaboratoireTotal / compteur, 2))
            moyenneHeuresTotal = float(round(heureTotal / compteur, 2))
            moyennesPrescripteurs.append(moyenneHeuresPrescripteur)
            moyennesLaboratoire.append(moyenneHeuresLaboratoire)
            moyennesTotal.append(moyenneHeuresTotal)
            heurePrescripteurTotal = 0
            heureLaboratoireTotal = 0
            heureTotal = 0
            compteur = 0
            
        heurePrescripteurTotal += row['Delai Prescripteur']
        heureLaboratoireTotal += row['Delai Laboratoire']
        heureTotal += row['Delai Total']
        compteur += 1
        
        if(i == len(S)-1):
            moyenneHeuresPrescripteur = float(round(heurePrescripteurTotal / compteur, 2))
            moyenneHeuresLaboratoire = float(round(heureLaboratoireTotal / compteur, 2))
            moyenneHeuresTotal = float(round(heureTotal / compteur, 2))
            moyennesPrescripteurs.append(moyenneHeuresPrescripteur)
            moyennesLaboratoire.append(moyenneHeuresLaboratoire)
            moyennesTotal.append(moyenneHeuresTotal)
            
    return moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal



#Transforme la durée en heure en delta de temps
def dateVersTexte(moyennes):
    moyennesFormates = []
    for i in range(len(moyennes)):
        resultat = moyennes[i] * pd.Timedelta(hours=1)
        if(resultat.nanoseconds == 999):
            resultat += pd.Timedelta(nanoseconds=1)
        moyennesFormates.append(str(resultat))
    return moyennesFormates


#Renvoie une liste des dates
def date(S): 
    Date= []
    T=S['Date Prel'].values
    for elt in T:
        if elt not in Date :
            Date.append(elt)
    return Date

#Retourne les statistiques des résultats pour chaque date
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

#Retourne les statistiques des résultats pour chaque Structure
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
  

#Retourne les statistiques des résultats pour chaque Date-Structure
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

#Ajoute une ligne avec le total par colonne 
#Ajoute une colonne aevc le total par ligne
def addtot(R): 
    S=R.copy()
    tot=S.apply(np.sum, axis =0).values
    S.loc['TOTAL']=tot
    S['TOTAL']=S['POS']+S['NEG']+S['IND']+S['NON CONFORME']
    return S

def addTotalDelai(moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal):
    moyennesPrescripteurs.append(mean(moyennesPrescripteurs))
    moyennesLaboratoire.append(mean(moyennesLaboratoire))
    moyennesTotal.append(mean(moyennesTotal))
    return moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal
    
    

#Formate toutes les colonnes de statistiques avec le type int
def page1sepop(S): #stat pos neg pctge 
    col=['Date','Opération',
         'Delai moyen prescripteur','Delai moyen laboratoire','Delai moyen total',
         'Delai moyen prescripteur en heures','Delai moyen laboratoire en heures','Delai moyen total en heures',
         'Type', 'POS' , 'NEG' , 'IND' , 'NON CONFORME' , 'TOTAL','Salicov','Non reçu','Tube fuyant','Volume non respecté','Discordance','Tube vide','Prélèvement d\'expectoration','Contenant non adapté','Absence d\'identité','Autre']
    Q=S[col]
    for elt in col[9:23]:
        Q[elt]=Q[elt].astype(int) 
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


Tk().withdraw()

fichiercsv=askopenfilename(title = 'sélectionner extraction DOSSIER contenant les ARP')
folder = askdirectory(title = 'sélectionner dossier cible')
Tk().withdraw()
os.chdir(folder)


print("Extraction des donnéees du fichier CSV")
L=extractdoss(fichiercsv)

print("Retrait des tests")
Q=restdoss(L)
R=fuscoldoss(Q)

print("Génération des listes")
Presc=presc(R)
Datepresc=datepresc(R)
Date = date(R)

print("Génération des Moyennes des délais des prescripteurs et du laboratoire")
moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal = moyennes(R)

print("Ajout des valeurs moyennes pour la totalité des opérations")
moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal = addTotalDelai(moyennesPrescripteurs, moyennesLaboratoire, moyennesTotal)

print("Génération des délais sous forme de texte")
moyennesPrescripteursTexte = dateVersTexte(moyennesPrescripteurs)
moyennesLaboratoireTexte = dateVersTexte(moyennesLaboratoire)
moyennesTotalTexte = dateVersTexte(moyennesTotal)

print("Génération des statistiques des résultats pour chaque Structure")
Stat=stats(R,Presc)

print("Génération des statistiques des résultats pour chaque Date-Structure")
Statdatep=statsdatep(R,Datepresc)

print("Génération des statistiques des résultats pour chaque Date")
Statdate=statsdate(R,Date)

Statdate=Statdate.loc[[elt for elt in Statdate.index if 'TEST' not in elt.upper()]]
#restreint R aux indices covisan
Statdatep=Statdatep.loc[[elt for elt in Statdatep.index if 'TEST' not in elt.upper()]]
#Stat : contient les stats (voir colonnes)
SOPS=Statdatep.loc[[elt for elt in Statdatep.index]]

DFOPS=addtot(SOPS)
INDOP=DFOPS.index.values
DFOPS['Date']=[elt[:10] for elt in INDOP]
DFOPS['Opération']=[elt[12:] for elt in INDOP]
DFOPS['Delai moyen prescripteur en heures'] = moyennesPrescripteurs
DFOPS['Delai moyen laboratoire en heures'] = moyennesLaboratoire
DFOPS['Delai moyen total en heures'] = moyennesTotal
DFOPS['Delai moyen prescripteur'] = moyennesPrescripteursTexte
DFOPS['Delai moyen laboratoire'] = moyennesLaboratoireTexte
DFOPS['Delai moyen total'] = moyennesTotalTexte
DFOPS['Type']='PCR'


DFOPS=page1sepop(DFOPS)
expnoind(DFOPS,generer_date()+' Bilan OPS pcr.xlsx')
