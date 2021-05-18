import pandas as pd
import csv
import io
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.simpledialog import askstring
from tkinter.simpledialog import askinteger


Tk().withdraw() 
    
"changer nom orga"
BRS='BRS'
TROD='TROD'
vislvl=4

fichier_source=askopenfilename(title = 'Fichier avec infos structure')
folder = askdirectory(title = 'dossier cible')
#tel1=askstring(title = 'téléphone préleveur' , prompt = 'entrer le numéro de téléphone du préleveur')
tel1 = '0610101010'
dateetab=askstring(title = 'nom_fichier' , prompt = 'format _AAAAMMJJ_OP')
NB=askinteger(title ='Avec/Sans PCR jumelle', prompt = '1 : uniquement TROD\n2 :TROD + PCR')
os.chdir(folder)    


def extractioncsv(fichiercsv):  
    liste = [] 
    with open(fichiercsv, encoding='cp1252') as fcsv :
        lecteur = csv.reader(fcsv, delimiter=';') 
        for ligne in lecteur: 
            liste.append(ligne) 
        return liste

def extraction(fichierxls):
    L=[]
    xls=pd.ExcelFile(fichierxls)
    sheets=xls.sheet_names
    for elt in sheets:
        L.append(pd.read_excel(xls , elt))
    return L,sheets


def cpostal(L,S):
    cp='CODE POSTAL'
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    return scp

def FIN(L,S):
    cp='N° FINESS'
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    return scp     
        
def spe(L,S):
    cp='SPECIALITE'
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    return scp   

def nometa(L,S):
    cp='NOM ETABLISSEMENT'
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    return scp         

def ville(L,S):
    cp='VILLE'
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    return scp   

def exportcsv(V,filename):
    s=io.StringIO(V)
    with open(filename,'w+') as file:
        for line in s: 
            file.write(line)


def nbcar(i):
    if i<10:
        return 1
    else : 
        return 1 + nbcar(i//10)


def groupe(FINESS,CP,SPE,NOM,VIL,NB):
    g=''
    if NB==2:
        g+=';'+FINESS+';'+SPE+' '+NOM+' '+VIL+';F;'+BRS+';;;;'+CP+';'+VIL
    return g


def presc(FINESS,CP,SPE,NOM,VIL,NB):
    p=''
    p+='PTROD_BRS_'+FINESS+';;MED_TROD_'+SPE+' '+NOM+' '+VIL+';;'+TROD+';;;;;;;;;;;;;'
    if NB==2:
        p+='\n'
        p+='P_'+FINESS+'_PRO;;MED_'+SPE+' '+NOM+' '+VIL+';;'+BRS+';;;;;;;;;;;;;'
    
    
    gp=''
    gp+='PTROD_BRS_'+FINESS+';750100141_TROD'+';'+TROD
    if NB==2:
        gp+='\n'
        gp+='P_'+FINESS+'_PRO;'+FINESS+';'+BRS
    
    
    
    return p,gp

def acc(FINESS,CP,SPE,NOM,VIL,NB,tel1):
    ac=''
    ac+='TRD_PRL_'+(SPE+NOM+VIL).replace(' ','')+';PRELEVEUR TROD '+SPE +' '+NOM+' '+VIL+';;;PTROD_BRS_'+FINESS+';2_PRELEVEUR_TROD;FR;FR;YES;Cyber2020=;;2_PRELEVEUR_TROD;NO;;;'+tel1 +';;;;;YES;'+TROD 
    if NB==2:
        ac+='\n'
        ac+='BRS_PRL_'+(SPE+NOM+VIL).replace(' ','')+';PRELEVEUR PCR '+SPE +' '+NOM+' '+VIL+';;;P_'+FINESS+'_PRO;1_PRELEVEUR;FR;FR;YES;Cyber2020=;;1_PRELEVEUR;NO;;;'+tel1 +';;;;;YES;'+BRS
    
    ga=''
    ga+='TRD_PRL_'+(SPE+NOM+VIL).replace(' ','')+';750100141_TROD'+';'+TROD
    if NB==2:
        ga+='\n'
        ga+='BRS_PRL_'+(SPE+NOM+VIL).replace(' ','')+';'+FINESS+';'+BRS
    
    return ac,ga



def informations(FINESS,CP,SPE,NOM,VIL,NB,tel1):

    
    if NB==2:
        text1=open('infos GLIMS.txt','w')
        gstring=CP + '\n'
        gstring+='Groupe : '+ FINESS 
        gstring+='\nCode : ' + 'P_'+FINESS+'_PRO \nLibellé : MED PRO ' + SPE + ' '+ NOM + ' '+ VIL
        text1.write(gstring)
        text1.close
    
    
    text2=open('infos comptes.txt','w')
    tstring=''
    tstring += 'prel trod login : ' + 'TRD_PRL_'+(SPE+NOM+VIL).replace(' ','')+'    mdp : Cyber2020=  tel : '+tel1
    if NB==2:
        tstring+='\n'
        tstring+='prel pcr login : ' 'BRS_PRL_'+(SPE+NOM+VIL).replace(' ','')+'    mdp : Cyber2020=  tel : '+tel1

    text2.write(tstring)
    text2.close()

    
os.chdir(folder)
L,S = extraction(fichier_source)
    
CP=cpostal(L,S)
FINESS=FIN(L,S)
NOM=nometa(L,S)
SPE=spe(L,S)
VIL=ville(L,S)

g=groupe(FINESS,CP,SPE,NOM,VIL,NB)
i,gi=presc(FINESS,CP,SPE,NOM,VIL,NB)
ac,ga=acc(FINESS,CP,SPE,NOM,VIL,NB,tel1)

exportcsv(g, '1_Groups'+ dateetab + '.csv')
exportcsv(i,'2_1_Issuer'+ dateetab + '.csv')
exportcsv(gi,'2_2_Group_Issuer'+ dateetab + '.csv')
exportcsv(ac,'3_1_Accounts'+ dateetab + '.csv')
exportcsv(ga,'3_2_Group_Accounts'+ dateetab + '.csv')

informations(FINESS,CP,SPE,NOM,VIL,NB,tel1)