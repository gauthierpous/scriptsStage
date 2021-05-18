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
    
patientcodingsystem='BRS'

vislvl=4





fichier_pers=askopenfilename(title = 'Fichier avec infos structure')
folder = askdirectory(title = 'dossier cible')
#tel1=askstring(title = 'téléphone prescripteur' , prompt = 'entrer le numéro de téléphone du prescripteur')
tel1 = '0610101010'
#tel2=askstring(title = 'téléphone préleveur' , prompt = 'entrer le numéro de téléphone du préleveur')
tel2 = '0610101010'
dateetab=askstring(title = 'nom_fichier' , prompt = 'format _AAAAMMJJ_NOMETABLISSEMENT')

nbreacc=0
while nbreacc != 1 and nbreacc != 2:
    nbreacc=askinteger(title ='combien de status de testés à créer?', prompt = 'entrer 1 si uniquement soignants/pros \n entrer 2 si soignant/pros et résidents')
#nbreacc = 1 ou 2 en fonction du nb de comtpes à créer.
    
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




def genererCsvGroupe(L,S):
    infos = ['NOM ETABLISSEMENT', 'N° FINESS','SPECIALITE', 'ADRESSE 1', 'ADRESSE 2', 'CODE POSTAL', 'VILLE', 'TEL. ETABLISSEMENT', 'FAX ETABLISSEMENT', 'E-MAIL ETABLISSEMENT', 'GENRE', 'NOM DIRECTEUR', 'PRENOM DIRECTEUR', 'DATE DE NAISSANCE', 'METIER', 'CODE UH RATTACHEMENT', 'TEL. PORT. DIRECTEUR', 'FAX DIRECTEUR', 'E-MAIL DIRECTEUR']
    V=""
    for elt in infos:
        for i, titre in enumerate(L[0].values[:,1]):
            if titre ==elt:
                if (str(L[0].values[i,2]) != 'nan'):
                    V+=str(L[0].values[i,2])
                V+=';'
                break
    
    return V[:-1]



def genererCsvResidents(L,S):
    infos = ['NOM DE NAISSANCE' , 'NOM USUEL' , 'PRENOM' , 'SEXE', 'DATE DE NAISSANCE' ]
    if 'PRENOM' in L[1].columns:
        Newdf=L[1]
    else :
        i=-1
        Tab=L[1].values #le tableau ave les résidents
        stop=False
        while not stop:
            i+=1
            ligne = Tab[:,i]
            for elt in ligne:
                if elt == 'PRENOM':
                    stop = True
                    break
        
        
        """i = indice des labels de colonnes"""
        NewTab=Tab[i:,]
        Lab=Tab[i-1,]
        Newdf = pd.DataFrame(NewTab,columns=Lab)
    V=''
    mustbreak=False
    for i, row in Newdf.iterrows():
        for elt in infos:
            if 'NaT'==str(row[elt]):
                mustbreak=True
                break
            if str(row[elt]) != 'nan':
                V+= str(row[elt])
            V += ';'  #retire les ; fin de ligne
        V=V[:-1]
        if mustbreak:
            break
        V+='\n'
        
    return V
        
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




def generermanoresidents(L,S,idepart,patientcodingsystem, codeville, visibilitylevel):
    infos = ['NOM DE NAISSANCE' , 'PRENOM' , 'NOM USUEL','ESP', 'DATE DE NAISSANCE' , 'SEXE', 'VERIFICATION STATUS' , 'TEL. PORTABLE' , 'E-MAIL Personnel' , 'CONS' , 'DECEASED' ]
    if 'PRENOM' in L[1].columns:
        Newdf=L[1]
        Lab=L[1].columns
    else :
        i=-1
        Tab=L[1].values #le tableau ave les résidents
        stop=False
        while not stop:
            i+=1
            ligne = Tab[:,i]
            for elt in ligne:
                if elt == 'PRENOM':
                    stop = True
                    break
        
        
        """i = indice des labels de colonnes"""
        NewTab=Tab[i:,]
        Lab=Tab[i-1,]
        Newdf = pd.DataFrame(NewTab,columns=Lab)
    V=''
    mustbreak=False
    j=0 #comptage pour changer le idepart
    for i, row in Newdf.iterrows():
        j+=1
        if 'NaT'==str(row['NOM DE NAISSANCE']):
            mustbreak = True
        if mustbreak:
            break
        V+= patientcodingsystem  #code à générer
        V+=';' + codeville #
        V+=(10-nbcar(i+idepart)) * '0'
        V+=str(i+   idepart)
        V+=';'+ str(visibilitylevel) + ';' 
        for elt in infos:
            if elt in Lab:
                if str(row[elt]) != 'nan': #pour ne pas afficher 'nan'
                    if elt == 'E-MAIL Personnel':
                        V+=str(row[elt]).lower()
                    elif elt == 'DATE DE NAISSANCE':
                        V+=str(row[elt])[0:4] + str(row[elt])[5:7] + str(row[elt])[8:10]
                    elif elt == 'TEL. PORTABLE':
                        if str(row[elt])[0]!='0':
                            V+='0'
                        V+=(str(row[elt]).replace(' ','')).replace('.','')
                        
                    else:
                        V+= str(row[elt])
            if elt == 'VERIFICATION STATUS':
                V+='T'
            if elt == 'CONS':
                V+='T'
            if elt == 'DECEASED':
                V+='F'    
                
                
            V += ';'  
        V=V[:-1]  #retire les ; fin de ligne
        V+='\n'
        
    return V,j
        


def generermanogroupe(L,S,patientcodingsystem):
    infos = ['ESP','N° FINESS','SPECIALITE','NOM ETABLISSEMENT', 'VILLE' ,'F', patientcodingsystem , 'ESP','ESP','ESP', 'CODE POSTAL','VILLE','ESP','ESP','ESP','ESP','ESP','ESP','ESP']
    V=""
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        if elt == patientcodingsystem:
            V+=patientcodingsystem
            V+=';'
        if elt == 'F':
            V+='F'
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
            
    
        if elt  == 'SPECIALITE':
            V=V[:-1]
            V+=' '
        if elt  =='NOM ETABLISSEMENT':
            V=V[:-1]
            V+=' -'
    
    return V[:-1]

def generermanoissuer(L,S):
    infos = ['P_' , 'N° FINESS','_PRO','ESP' ,'MED SOIGNANT - ','SPECIALITE','NOM ETABLISSEMENT', 'VILLE' ,'ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','CODE POSTAL','VILLE']
    V=""
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
        if elt  == 'SPECIALITE':
            V=V[:-1]
            V+=' '
        if elt  =='NOM ETABLISSEMENT':
            V=V[:-1]
            V+=' -'
        if elt  == 'P_':
            V+='P_'
        if elt  == 'N° FINESS':
            V=V[:-1]      
        if elt  == 'MED SOIGNANT - ':
            V+='MED PRO - ' 
        if elt == '_PRO':
            V+='_PRO'
            V+=';'
        
    V=V[:-1]
    
    if nbreacc==1:
        return V
    V+='\n'    
         
    infos = ['P_' , 'N° FINESS','_RES','ESP' ,'MED RESIDENT - ','SPECIALITE','NOM ETABLISSEMENT', 'VILLE' ,'ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','ESP','CODE POSTAL','VILLE' ]
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
        if elt  == 'SPECIALITE':
            V=V[:-1]
            V+=' '
        if elt  =='NOM ETABLISSEMENT':
            V=V[:-1]
            V+=' -'
        if elt  == 'P_':
            V+='P_'
        if elt  == 'N° FINESS':
            V=V[:-1]      
        if elt  == 'MED RESIDENT - ':
            V+='MED RESIDENT   - ' 
        if elt == '_RES':
            V+='_RES'
            V+=';'

        
    return V[:-1]


def generermanogroupissuer(L,S,patientcodingsystem):
    infos = ['P_' , 'N° FINESS','_PRO','N° FINESS',patientcodingsystem]
    V=""
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        if elt == patientcodingsystem:
            V+= patientcodingsystem
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
        if elt  == 'P_':
            V+='P_'
 
        if elt == '_PRO':
            V=V[:-1]
            V+='_PRO'
            V+=';'
        
    V=V[:-1]
    
    if nbreacc==1:
        return V
    
    V+='\n'
    
    
    infos = ['P_' , 'N° FINESS','_RES','N° FINESS',patientcodingsystem]
    
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        if elt == patientcodingsystem:
            V+= patientcodingsystem
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
        if elt  == 'P_':
            V+='P_'      
        if elt == '_RES':
            V=V[:-1]
            V+='_RES'
            V+=';'    
        
            
            
            
    return V[:-1]


def genereracountsgroupaccounts(L,S,patientcodingsystem,tel1,tel2):
    A=""
    G=""
    NOM=""  #va contenir nom login
    IN="NOM ETABLISSEMENT"

    IV="VILLE"
    
    info= 'N° FINESS'
    for i, titre in enumerate(L[0].values[:,1]):
        if titre == IN:
            NOM+= str(L[0].values[i,2])
        if titre == IV:
            NOM+= str(L[0].values[i,2])
        
        if titre == info:
            siret=str(L[0].values[i,2]) #on vient chercher le siret
    print(NOM)
    NOM=NOM.replace(" ",'')
    NOM=NOM.replace('nan','')
    
    if nbreacc==1:
        G+="BRS_PRS_" +NOM+ 'PRO;' + siret + ';'+  patientcodingsystem
        G+='\n'
        G+="BRS_PRL_" +NOM+ 'PRO;' + siret + ';'+  patientcodingsystem
        
    
    elif nbreacc==2:
        G+="BRS_PRS_" +NOM+ 'PRO;'+ siret + ';'+ patientcodingsystem
        G+='\n'
        G+="BRS_PRS_" +NOM+ 'RES;'+ siret + ';'+  patientcodingsystem
        G+='\n'
        G+="BRS_PRL_" +NOM+ 'PRO;' + siret + ';'+  patientcodingsystem
        G+='\n'
        G+="BRS_PRL_" +NOM+ 'RES;' + siret + ';'+ patientcodingsystem
    
    infos = ['SPECIALITE','NOM ETABLISSEMENT', 'VILLE']
    T=""

    
    for elt in infos:
        for i, titre in enumerate(L[0].values[:,1]):
            if titre ==elt:
                if (str(L[0].values[i,2]) != 'nan'):
                    T+=str(L[0].values[i,2]).upper()
                T+=';'
                break
        if elt == 'SPECIALITE':
            T=T[:-1]
            T+=' ' 
        if elt == 'NOM ETABLISSEMENT':  
            T=T[:-1]
            T+=' - '
    
    #T contient l'ensemble spe etabl ville, sans ; à la fin
    if nbreacc==1:
        A= "BRS_PRS_"  +NOM +  'PRO;MEDECIN Pros. ' + T + ';;P_' + siret +'_PRO;1_PRESCRIPTEUR;FR;FR;YES;Cyber2020=;;1_PRESCRIPTEUR;NO;;;' + tel1 + ';;;;;YES;' + patientcodingsystem
        A+='\n'
        A+="BRS_PRL_"  +NOM +  'PRO;PRELEVEUR pros. ' + T + ';;P_' + siret +'_PRO;1_PRELEVEUR;FR;FR;YES;Cyber2020=;;1_PRELEVEUR;NO;;;' + tel2 + ';;;;;YES;' + patientcodingsystem
    
    
    
    
    else :
        A= "BRS_PRS_"  +NOM + 'PRO;MEDECIN Pros. ' + T + ';;P_' + siret +'_PRO;1_PRESCRIPTEUR;FR;FR;YES;Cyber2020=;;1_PRESCRIPTEUR;NO;;;' + tel1 + ';;;;;YES;' + patientcodingsystem
        A+='\n'
        A+="BRS_PRS_"  +NOM + 'RES;MEDECIN Res. ' + T + ';;P_' + siret +'_RES;1_PRESCRIPTEUR;FR;FR;YES;Cyber2020=;;1_PRESCRIPTEUR;NO;;;' + tel1 + ';;;;;YES;' + patientcodingsystem
        A+='\n'
        A+="BRS_PRL_"  +NOM + 'PRO;PRELEVEUR Pros. ' + T + ';;P_' + siret +'_PRO;1_PRELEVEUR;FR;FR;YES;Cyber2020=;;1_PRELEVEUR;NO;;;' + tel2 + ';;;;;YES;' + patientcodingsystem
        A+='\n'
        A+="BRS_PRL_"  +NOM + 'RES;PRELEVEUR Res. ' + T + ';;P_' + siret +'_RES;1_PRELEVEUR;FR;FR;YES;Cyber2020=;;1_PRELEVEUR;NO;;;' + tel2 + ';;;;;YES;' + patientcodingsystem
    
    return A,G,NOM


def informations(L,NOM):
    cp='CODE POSTAL'
    lib=['SPECIALITE','NOM ETABLISSEMENT', 'VILLE']
    V=''
    for elt in lib:
        for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
            if titre ==elt: #si ça match :
                if (str(L[0].values[i,2]) != 'nan'): #si non vide
                    V+=str(L[0].values[i,2]).upper()
                    V+=' '
                    break
    scp=''
    for i,titre in enumerate(L[0].values[:,1]):
        if titre == cp:
            scp+=str(L[0].values[i,2])
            break
    
    text1=open('infos GLIMS.txt','w')
    for i, titre in enumerate(L[0].values[:,1]):
        if titre == 'N° FINESS':
            siret=str(L[0].values[i,2])
    #siret: le siret/finess de l'et
    tstring=scp +'\n'
    tstring+='Group : ' + siret +'\n' + 'Issuer :'
    tstring+='\nCode : ' + 'P_'+siret+'_PRO \nLibellé : MED PRO - ' + V
    if nbreacc==2:
        tstring+='\nCode : ' + 'P_'+siret+'_RES \nLibellé : MED RES - ' + V
    text1.write(tstring)
    text1.close
    
    
    text2=open('infos comptes.txt','w')
    tstring=''
    if nbreacc==1:
        tstring+='Prescripteur login : ' + 'BRS_PRS_' +NOM + 'PRO   Mdp : Cyber2020=  tel: ' + tel1 +'\n'
        tstring+='Préleveur login : ' + 'BRS_PRL_' +NOM + 'PRO  Mdp : Cyber2020=  tel: ' + tel2 
    
    if nbreacc==2:
        tstring+='Prescripteur Pro login : ' + 'BRS_PRS_'+NOM +'PRO  Mdp : Cyber2020=  tel: ' + tel1 +'\n'
        tstring+='Prescripteur Res login : ' +'BRS_PRS_'+NOM + 'RES  Mdp : Cyber2020=  tel: ' + tel2 +'\n'
        tstring+='Préleveur Pro login : ' + 'BRS_PRL_' +NOM+ 'PRO  Mdp : Cyber2020=  tel: ' + tel1 +'\n'
        tstring+='Préleveur Res login : ' + 'BRS_PRL_'+NOM+ 'RES  Mdp : Cyber2020=  tel: ' + tel2

    text2.write(tstring)
    text2.close()
    if len(scp) !=5 :
        print('!code potal absent/invalide!')

def genererlogin(L,S):
    infos = ['MED SOIGNANT - ','SPECIALITE','NOM ETABLISSEMENT', 'VILLE' ]
    V=""
    for elt in infos: #on cherche une certaine ligne
        if elt == 'ESP':
            V+=';'
        else :
            for i, titre in enumerate(L[0].values[:,1]): #parcours 2e colonne ou ya les labels (par ligne ici)
                if titre ==elt: #si ça match :
                    if (str(L[0].values[i,2]) != 'nan'): #si non vide
                        V+=str(L[0].values[i,2]).upper()
                    V+=';'
                    break
        if elt  == 'SPECIALITE':
            V=V[:-1]
            V+=' '
        if elt  =='NOM ETABLISSEMENT':
            V=V[:-1]
            V+=' -'
        if elt  == 'P_':
            V+='P_'
        if elt  == 'N° FINESS':
            V=V[:-1]      
        if elt  == 'MED SOIGNANT - ':
            V+='MED PRO - ' 
        if elt == '_PRO':
            V+='_PRO'
            V+=';'
        #print(V)
    V=V[:-1]
    
    if nbreacc==1:
        return V
    V+='\n'    
        
    return V[:-1]



os.chdir(folder)    


L,S = extraction(fichier_pers)





I=generermanoissuer(L,S)

GpI=generermanogroupissuer(L,S,patientcodingsystem)

Gp=generermanogroupe(L,S,patientcodingsystem)

A,G,NOM = genereracountsgroupaccounts(L,S,patientcodingsystem,tel1,tel2)

log=genererlogin(L, S)
login=log+";"+"BRS_PRL_"+NOM+"PRO"+"\n"

exportcsv(Gp, '1_Groups'+ dateetab + '.csv')
exportcsv(I,'2_1_Issuer'+ dateetab + '.csv')
exportcsv(GpI,'2_2_Group_Issuer'+ dateetab + '.csv')
exportcsv(A,'3_1_Accounts'+ dateetab + '.csv')
exportcsv(G,'3_2_Group_Accounts'+ dateetab + '.csv')
informations(L,NOM)

file=open('C:/Users/4165306/Desktop/script/script/document.csv','a')
file.write(login)
file.close()
