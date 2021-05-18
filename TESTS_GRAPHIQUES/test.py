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
    infdoss = ['N° Patient BROUSSAIS' , 'Date de naissance' , 'Sexe du patient','Code postal du patient' ,'Date de prélèvement de dossier', 'Nom du prescripteur' ,  '94500-6','COV19_ARN_RR_PR_THF1_BPO' ,   'APSYM','XXX_DELAI_SYMPT']
    Q=L.rename(columns = {'N° patient Ajaccio' :'N° Patient BROUSSAIS' })
    Q=Q.query('Nom != "Patient_Test_11"')
    Q=Q.query('Nom != "Patient_Test_31"')
    Q=Q.query('Nom != "001_FORMATION_NOM"')
    Q=Q.query('Nom != "Patient_Test_20"')
    Q=Q.query('Nom != "Patient_Advens"')
    Q=Q.query('Nom != "Patient_Test_30_modifGlims"')
    Q=Q.query('Nom != "Patient_Test_12"')
    Q=Q.query('Nom != "Patient_Test_10"')
    Q=Q.query('Nom != "Patient_Test_33"')
    Q=Q.query('Nom != "TEST TEST"')
    Q=Q.query('Nom != "TEST_mail"')
    Q=Q.query('Nom != "TEST_RPPS_traitant"')
        
    Q=Q[['N° Patient BROUSSAIS' , 'Date de naissance' , 'Sexe du patient','Code postal du patient' ,'Date de prélèvement de dossier', 'Nom du prescripteur' ,  '94500-6',   'APSYM','Heure de prélèvement de dossier','Date_CR','Heure du dernier compte-rendu de résultat']]
    Q=Q.rename(columns = {'Date de prélèvement de dossier' : 'Date de prélèvement'})
    Q=Q.rename(columns = {'Nom du prescripteur' : 'Prescripteur'})
    Q=Q.rename(columns = {'Heure de prélèvement de dossier' : 'Heure de prélèvement'})
    Q=Q.rename(columns = {'Heure du dernier compte-rendu de résultat' : 'Heure du CR'})
    Q=Q.query('Prescripteur != "APHP-HUPNVS"')
    Q=Q.query('Prescripteur != "APHP-HUPC"')
    Q=Q.query('Prescripteur != "ACT EXT IVG"')
    return Q
    

def fuscoldoss(Q):
    n=len(Q)
    ind=[i for i in range(n)]
    Q.index=ind
    Q['Res_test']=0
    Q['Symp']=0
    Q=Q.fillna(0)
    C=['Prescripteur' , 'Date de naissance' , 'Sexe du patient','Code postal du patient' , 'Date de prélèvement' , 'Res_test' , 'Symp','Heure de prélèvement','Date_CR','Heure du CR']
    nc = len(C)
    r=np.empty((n,nc),dtype=object)
    for i,row in Q.iterrows():
        r[i,0]=row['Prescripteur']
        r[i,1]=row['Date de naissance']
        r[i,2]=row['Sexe du patient']
        r[i,3]=row['Code postal du patient']
        r[i,4]=row['Date de prélèvement']
        r[i,7]=row['Heure de prélèvement']  
        r[i,8]=row['Date_CR']  
        r[i,9]=row['Heure du CR']          
        v94=Q.loc[i,'94500-6']
        if v94 in ['POSITIF', 'positif','Positif','P']:
            r[i,5]='POS'
        elif v94 in ['*Négatif','Négatif' , 'négatif','N',]:
            r[i,5]='NEG'
        elif v94 in ['I','indeterminé','Indeterminé','Indéterminé']:
            r[i,5]='IND'
        elif v94 in ['prélèvement non conforme','Prélèvement non conforme']:
            r[i,5]='NCONF'
        vsymp=Q.loc[i,'APSYM']
        if vsymp == 'Asymptomatique':
            r[i,6]='ASYMPTOMATIQUE'
        elif vsymp == 'Ne sait pas':
            r[i,6]='INCONNU'
        else:
            r[i,6]='SYMPTOMATIQUE'
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
        dp=row['Date de prélèvement']+' - '+row['Prescripteur']
        if dp not in Presc:
            if 'TEST' not in dp.upper():
                Presc.append(dp)
    return Presc
        

def date(S):
    Date= []
    T=S['Date de prélèvement'].values
    for elt in T:
        if elt not in Date :
            Date.append(elt)
    return Date
        
def statsdate(S,Date):
    C=[ 'POS', 'POSYMPTO', 'POSASYMPTO', 'POSINCONNU', 'NEG', 'NEGSYMPTO', 'NEGASYMPTO','NEGINCONNU','IND','INDSYMPTO', 'INDASYMPTO','INDINCONNU','NON CONFORME']            
    Index= Date
    Np=len(Date)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    for i,row in S.iterrows():
        ind=row['Date de prélèvement']
        test=row['Res_test']
        symp=row['Symp']
        if ind in Date:
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['POSASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['POSYMPTO']+=1
                else :
                    D.loc[ind]['POSINCONNU']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['NEGASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['NEGSYMPTO']+=1
                else :
                    D.loc[ind]['NEGINCONNU']+=1
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['INDASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['INDSYMPTO']+=1
                else :
                    D.loc[ind]['INDINCONNU']+=1 
    return D

def stats(S,Presc):
    C=[ 'POS', 'POSYMPTO', 'POSASYMPTO', 'POSINCONNU', 'NEG', 'NEGSYMPTO', 'NEGASYMPTO','NEGINCONNU','IND','INDSYMPTO', 'INDASYMPTO','INDINCONNU','NON CONFORME']            
    Index= Presc
    Np=len(Presc)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    for i,row in S.iterrows():
        ind=row['Prescripteur']
        test=row['Res_test']
        symp=row['Symp']
        if ind in Presc:
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['POSASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['POSYMPTO']+=1
                else :
                    D.loc[ind]['POSINCONNU']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['NEGASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['NEGSYMPTO']+=1
                else :
                    D.loc[ind]['NEGINCONNU']+=1
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['INDASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['INDSYMPTO']+=1
                else :
                    D.loc[ind]['INDINCONNU']+=1 
    return D



def calculate_age(dtob):
    today = Date1.today()
    return today.year - dtob.year - ((today.month, today.day) < (dtob.month, dtob.day))    
    
def sec_to_hours(seconds):
    a=str(seconds//3600)
    b=str((seconds%3600)//60)
    c=str((seconds%3600)%60)
    d=["{} hours {} mins {} seconds".format(a, b, c)]
    return d


def statsdatep(S,Presc):
    C=[ 'POS', 'POSYMPTO', 'POSASYMPTO', 'POSINCONNU', 'NEG', 'NEGSYMPTO', 'NEGASYMPTO','NEGINCONNU','IND','INDSYMPTO', 'INDASYMPTO','INDINCONNU','NON CONFORME','HOMME','FEMME','AUTRE','0-10','11-20','21-30','31-40','41-50','50+','moyenne_neg','heures_neg','minutes_neg','secondes_neg','total_neg','moyenne_pos','heures_pos','minutes_pos','secondes_pos','total_pos']  
    Index= Presc
    Np=len(Presc)
    nc=len(C)
    D=pd.DataFrame(np.zeros((Np,nc)),columns=C,index=Index)
    D=D.astype(int)
    #print(S['Res_test'])
    #total_pos=0
    #total_neg=0
    for i,row in S.iterrows():
        ind=row['Date de prélèvement']+' - '+row['Prescripteur']
        test=row['Res_test']
        symp=row['Symp']
        sexe=row['Sexe du patient']
        birth=calculate_age(Date1(int(row['Date de naissance'][6:10]), int(row['Date de naissance'][3:5]), int(row['Date de naissance'][0:2])))
        date_prel=datetime(int(row['Date de prélèvement'][6:10]), int(row['Date de prélèvement'][3:5]), int(row['Date de prélèvement'][0:2]), int(row['Heure de prélèvement'][0:2]), int(row['Heure de prélèvement'][3:5]), 0)
        date_CR=datetime(int(row['Date_CR'][6:10]), int(row['Date_CR'][3:5]), int(row['Date_CR'][0:2]), int(row['Heure du CR'][0:2]), int(row['Heure du CR'][3:5]), 0)
        delai_CR=date_CR-date_prel
        if ind in Presc:
            total_seconds_pos=0
            total_seconds_neg=0
            if test == 'NCONF':
                D.loc[ind]['NON CONFORME']+=1
            elif test == 'POS':   #pos
                D.loc[ind]['POS']+=1
                total_seconds_pos+=int(delai_CR.total_seconds())
                #total_pos+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['POSASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['POSYMPTO']+=1
                else :
                    D.loc[ind]['POSINCONNU']+=1
            elif test == 'NEG':   #neg
                D.loc[ind]['NEG']+=1
                total_seconds_neg+=int(delai_CR.total_seconds())
                #total_neg+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['NEGASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['NEGSYMPTO']+=1
                else :
                    D.loc[ind]['NEGINCONNU']+=1
                    
            elif test == 'IND':   #ind
                D.loc[ind]['IND']+=1
                if symp=='ASYMPTOMATIQUE':
                    D.loc[ind]['INDASYMPTO']+=1
                elif symp == 'SYMPTOMATIQUE':
                    D.loc[ind]['INDSYMPTO']+=1
                else :
                    D.loc[ind]['INDINCONNU']+=1
            if sexe == "M":
                D.loc[ind]["HOMME"]+=1
            elif sexe == "F":
                D.loc[ind]["FEMME"]+=1
            else :
                D.loc[ind]["AUTRE"]+=1
            if birth < 10:
                D.loc[ind]["0-10"]+=1
            elif birth < 20:
                D.loc[ind]["11-20"]+=1
            elif birth < 30:
                D.loc[ind]["21-30"]+=1
            elif birth < 40:
                D.loc[ind]["31-40"]+=1
            elif birth < 50:
                D.loc[ind]["41-50"]+=1
            else:
                D.loc[ind]["50+"]+=1
            """print("bilan")
            print(delai_CR)
            print(int(delai_CR.total_seconds()))
            print(test)
            print(total_seconds_pos)"""
            #print(ind)
            #print(D.loc[ind]['POS'])
            #print("pos"+str(total_pos))
            if total_seconds_neg !=0:
                D.loc[ind]['total_neg']+=total_seconds_neg
                D.loc[ind]['moyenne_neg']=D.loc[ind]['total_neg']/D.loc[ind]['NEG']
                D.loc[ind]["heures_neg"]=int(D.loc[ind]['moyenne_neg']/3600)
                D.loc[ind]["minutes_neg"]=int((D.loc[ind]['moyenne_neg']%3600)/60)
                D.loc[ind]["secondes_neg"]=int((D.loc[ind]['moyenne_neg']%3600)%60)
            if total_seconds_pos !=0:
                D.loc[ind]['total_pos']+=total_seconds_pos
                D.loc[ind]['moyenne_pos']=D.loc[ind]['total_pos']/D.loc[ind]['POS']
                D.loc[ind]["heures_pos"]=int(D.loc[ind]['moyenne_pos']/3600)
                D.loc[ind]["minutes_pos"]=int((D.loc[ind]['moyenne_pos']%3600)/60)
                D.loc[ind]["secondes_pos"]=int((D.loc[ind]['moyenne_pos']%3600)%60)
                
    #print(D.loc[ind])
    
    #print(sec_to_hours(D.loc[ind]['total_secondes']))
    #print(D.loc[ind])
    return D
    


def addtot(R): #ajoute une ligne avec le total par colonne et par ligne
    S=R.copy()
    tot=S.apply(np.sum, axis =0).values
    S.loc['TOTAL']=tot
    S['TOTAL']=S['POS']+S['NEG']+S['IND']+S['NON CONFORME']
    return S
         
def page2(Q): #part sympto asympto
    S=Q.copy()
    c=['POSYMPTO' , 'POSASYMPTO' , 'POSINCONNU'] 
    for elt in c:
        S[elt]=S[elt].astype(str) +' (' + (round(10000*S[elt]/(S['POS'] + 0.0000000001))/100).astype(str) + '%)'
    c=[ 'NEGSYMPTO' , 'NEGASYMPTO' , 'NEGINCONNU' ]
    for elt in c:
        S[elt]=S[elt].astype(str) +' (' + (round(10000*S[elt]/(S['NEG'] + 0.0000000001))/100).astype(str) + '%)'
    c= ['INDSYMPTO' , 'INDASYMPTO' , 'INDINCONNU']
    for elt in c:
        S[elt]=S[elt].astype(str) +' (' + (round(10000*S[elt]/(S['IND'] + 0.0000000001))/100).astype(str) + '%)'

    return S
    
def page1(S): #stat pos neg pctge 
    col=['POS' , 'NEG' , 'IND' ,'NON CONFORME', 'TOTAL']
    Q=S[col]
    for elt in col[:4]:
        Q[elt]=Q[elt].astype(str) + ' (' + (round(10000*Q[elt]/Q['TOTAL'])/100).astype(str) +'%)'
    return Q

    
def page1sep(S): #stat pos neg pctge 
    col=['Date' , 'Presc' , 'POS' , 'NEG' , 'IND' , 'NON CONFORME' , 'TOTAL']
    Q=S[col]
    for elt in col[2:6]:
        Q[elt]=Q[elt].astype(str) + ' (' + (round(10000*Q[elt]/Q['TOTAL'])/100).astype(str) +'%)'
    return Q
    
def page1sepop(S): #stat pos neg pctge 
    col=['Date','Opération', 'POS' , 'NEG' , 'IND' , 'NON CONFORME' , 'TOTAL','HOMME','FEMME','AUTRE','0-10','11-20','21-30','31-40','41-50','50+','moyenne_neg','heures_neg','minutes_neg','secondes_neg','total_neg','moyenne_pos','heures_pos','minutes_pos','secondes_pos','total_pos']
    Q=S[col]
    for elt in col[2:16]:
        Q[elt]=Q[elt].astype(int) 
    return Q    
    


def expind(p1,filename):
    writer=pd.ExcelWriter(filename, engine = 'xlsxwriter')
    workbook=writer.book
    p1.to_excel(writer, sheet_name='Résultats')
    writer.save()                    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(folder +r"/"+ filename)
    ws1 = wb.Worksheets("Résultats")
    ws1.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()
    
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
    
def search_string_in_file(file_name, string_to_search):
    """Search for the given string in file and return lines containing that string,
    along with line numbers"""
    line_number = 0
    result=""
    # Open the file in read only mode
    with open(file_name, 'r') as read_obj:
        # Read all lines in the file one by one
        for line in read_obj:
            # For each line, check if line contains the string
            line_number += 1
            if string_to_search in line:
                # If yes, then add the line number & line as a tuple in the list
                result=line.rstrip()
    # Return list of tuples containing line numbers and lines where string is found
    return result


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






#export stats/date OK 
#Sdn=page1(addtot(Statdate))
#expind(Sdn,generer_date()+' Stats totales par date.xlsx')



#IND=Statdatep.index.values


#pour bilan ops
#SOPS=Statdatep.loc[[elt for elt in Statdatep.index if 'screening' not in elt and 'COVISAN' not in elt and 'VALIN' not in elt and 'Hôpital' not in elt and 'MAISON SANTÉ' not in elt and 'Broussais' not in elt and 'CH4V' not in elt and 'Abondances' not in elt and 'CH4V' not in elt and 'ROTHSCHILD' not in elt and 'HOPITAL' not in elt and 'TROUSSEAU' not in elt and 'SAMU' not in elt and 'SAINT-ANTOINE' not in elt and 'SAU NECKER' not in elt]]
SOPS=Statdatep.loc[[elt for elt in Statdatep.index]]
#

#Sdpn=addtot(Statdatep)  
#IND=Sdpn.index.values
#Sdpn['Date']=[elt[:10] for elt in IND]
#Sdpn['Presc']=[elt[13:] for elt in IND]

#Sdpn=page1sep(Sdpn)
#ok export total
#expnoind(Sdpn,generer_date()+' Stats totales par date et prescripteur.xlsx')


#Bilan OPS
DFOPS=addtot(SOPS)
INDOP=DFOPS.index.values
DFOPS['Date']=[elt[:10] for elt in INDOP]
DFOPS['Opération']=[elt[13:] for elt in INDOP]
"""DFOPS['Login']=""



for elt in INDOP :
    login=search_string_in_file('C:/Users/4165306/Desktop/script/script/document.csv', DFOPS['Opération'][elt])
    if login:
        print('Yes, string found in file')
        login=login.split(";")[1]
    else:
        print('String not found in file')
    if(elt != "TOTAL"):
        DFOPS['Login'][elt]=login"""
    
DFOPS=page1sepop(DFOPS)
expnoind(DFOPS,generer_date()+' Bilan OPS.xlsx')


