# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 10:37:08 2016

@author: user11
"""
import os
import shutil
import glob
import xlwings as xw
import numpy as np
import pandas as pd
from csv2xls import to_xls, file2df


path = r'G:\01-SUIVI PROCESSUS\GMM\Micro-Vu Vertex\Hirtenberger\27214161'
listp = ['P'+str(x) for x in range (1,51)]



listdf = []


# créer un dossier RR
try :
    os.mkdir(path+'/Cp')
except:
    pass

for P in listp:
    # lancer le programme vertex avec une ligne de commande
    fcsv = glob.glob(path+"/CSV Data/*.txt")[-1]  # File containing CSV
    data = file2df(fcsv)
    shutil.move(fcsv,path+"/Cp")
    df = data.copy()
    df['P'] = P #ajout du N° pièce
    # ajouter les résultats à la liste
    listdf.append(df)
df = pd.concat(listdf)
writer=pd.ExcelWriter(path+'/Cp/Cp.xlsx')
df.to_excel(writer,'sheet1')
writer.save()