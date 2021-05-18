import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.simpledialog import askinteger

Tk().withdraw()

#Demander le type de test à formater
#typeTest = 1, 2 ou 3 en fonction du type de test à formater.
typeTest = 0
codeType = "0"
while typeTest != 1 and typeTest != 2 and typeTest != 3 :
    typeTest = askinteger(title = 'Quel est le type de test à formater ?', prompt = 'Entrer 1 pour PCR \nEntrer 2 pour Salivaire \nEntrer 3 pour TROD')


#En fonction du chiffre reçu, affecter une valeur à codeType
if(typeTest == 1):
    codeType = "94500-6"
elif(typeTest == 2):
    codeType = "94845-5"
elif(typeTest == 3):
    codeType = "94558-4"



def extreq(filename):
    return pd.read_csv(filename,sep=',',header =1)



def split(L):
    if(typeTest == 1):
        Lres=L.query('Analyse == "94500-6"')
    elif(typeTest == 2):
        Lres=L.query('Analyse == "94845-5"')
    elif(typeTest == 3):
         Lres=L.query('Analyse == "94558-4"')
    Lsym=L.query('Analyse == "APSYM"')
    Lheb=L.query('Analyse == "TYPOR"')
    
    return Lres,Lsym,Lheb
    

def prep(Lres,Lsym,Lheb):
    Lres['ID']=Lres['Date de naissance']+Lres['Nom']+Lres['Prénom']+Lres['Prescripteur']
    Lsym['ID']=Lsym['Date de naissance']+Lsym['Nom']+Lsym['Prénom']+Lsym['Prescripteur']
    Lheb['ID']=Lheb['Date de naissance']+Lheb['Nom']+Lheb['Prénom']+Lheb['Prescripteur']
    Lsym=Lsym[['ID','Valeur']]
    Lheb=Lheb[['ID','Valeur']]
    return Lres,Lsym,Lheb

def stack(Qres,Qsym,Qheb):
    Q=Qres.merge(Qsym,left_on='ID',right_on='ID')
    Q=Q.merge(Qheb,left_on='ID',right_on='ID')
    return Q

def finaliser(Q):
    Q=Q.rename(columns = {'Valeur' : 'TYPOR'})
    Q=Q.rename(columns = {'Valeur_x' : codeType})
    Q=Q.rename(columns = {'Valeur_y' : 'APSYM'})
    Q=Q.rename(columns = {"Date du dernier compte-rendu de résultat" : 'Date_CR'})
    return Q[C]

def export(Q):
    Ldate=list(set(Q['Date_CR'].values))
    for elt in Ldate:
        Qd=Q.query('Date_CR == @elt')
        elt=elt.replace('/','-')
        Qd.to_csv(elt+'_BRS.csv',index = False,sep=';')
    


filename = askopenfilename(title = 'Export Résultat à Pousser')

folder = askdirectory(title = 'Dossier cible')

C=['Analyse', 'Demande', 'Prescripteur', 'Prénom du prescripteur','Nom du prescripteur', 'N° patient Ajaccio', 'Laboratoire',
       'Date de naissance', 'Code postal du patient', 'Nom', 'Prénom',
       'Lieu de naissance', 'Nom usuel', 'Deuxième prénom',
        codeType, 'TYPOR', 'APSYM',
       'Téléphone 1', 'Téléphone 2', 'Téléphone mobile', 'E-mail', 'Rue',
       'Numéro de maison', 'Complément au numéro de maison', 'Localité',
       'Pays', 'Langue', 'Sexe du patient', 'Date de prélèvement de dossier',
       'Heure de prélèvement de dossier', 
       'Date_CR',
       'Heure du dernier compte-rendu de résultat', 'Laboratoire exécutant',
 'Décédé(e)', 'Date de décès']




os.chdir(folder)

L=extreq(filename)
Lres,Lsym,Lheb=split(L)
Qres,Qsym,Qheb=prep(Lres,Lsym,Lheb)
Q=stack(Qres,Qsym,Qheb)
R=finaliser(Q)

export(R)
R.to_csv('BRS_dossier_compatible_bilanOPS.csv',sep=',')
