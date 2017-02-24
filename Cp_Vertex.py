# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 10:37:08 2016

@author: user11
"""
import os, sys
import shutil
import subprocess
import glob
import xlwings as xw
import numpy as np
import pandas as pd


path=os.getcwd()
sys.path.append(r"G:\01-SUIVI PROCESSUS\GMM\Micro-Vu Vertex\Python")
from csv2xls import to_xls, file2df


path = r'G:\01-SUIVI PROCESSUS\GMM\Micro-Vu Vertex\Schrader\43418-820'
cde = path + '\\iscmd.exe'
file = path + '/43418-820.iwp'
listp = ['P'+str(x) for x in range (1,51)]



listdf = []


# créer un dossier RR
try :
    os.mkdir(path+'/Cp')
except:
    pass

for P in listp:
    # lancer le programme vertex avec une ligne de commande
    subprocess.call("iscmd.exe /run" + path +"/43418-820.iwp")
    # Récupérer le fichier Csv
    fcsv = glob.glob(path+"/CSV Data/*.txt")[-1]  # File containing CSV
    # Transformation en Dataframe des données csv
    data = file2df(fcsv)
    # Déplacer le fichier dans un répertoire Cp
    shutil.move(fcsv,path+"/Cp")
    df = data.copy()
    df['P'] = P #ajout du N° pièce (Px)
    # ajouter les résultats à la liste
    listdf.append(df)
df = pd.concat(listdf)
writer=pd.ExcelWriter(path+'/Cp/Cp.xlsx')
df.to_excel(writer,'sheet1')
writer.save()
