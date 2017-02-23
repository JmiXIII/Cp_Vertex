# -*- coding: utf-8 -*-
"""
Created on Fri Feb 26 12:51:27 2016

@author: user11
"""

import os
import glob
import xlwings as xw
import numpy as np
import pandas as pd


color_OK = (208, 254, 182)
color_NOK = (236, 181, 167)


def file2df(fname):
    """CSV import function"""

    data = pd.read_csv(fname, header=None,
                       sep=None,
                       index_col=False,)
    return data
# %% Moulinette export xls


def to_xls(path):
    fcsv = glob.glob(path+"/CSV Data/*.txt")[-1]  # File containing CSV
    fxls = glob.glob(path+"/Layout/*.xlsx")[-1]  # xlsx file layout for report
    wb = xw.Workbook(fxls)
    data = file2df(fcsv)
    for index, row in data.iterrows():
        loc = data[0][index]
        value = data[1][index]
        if np.isnan(data[4][index]):
            if np.isnan(data[3][index]):
                color = color_OK
            else:
                color = color_NOK
        elif np.isnan(data[6][index]):
                    color = color_OK
        else:
            color = color_NOK
        xw.Range(str(loc)).value = value
        xw.Range(str(loc)).color = color
    os.system("pause")
    wb.xl_workbook.PrintOut()
    wb.close()
