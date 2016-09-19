# -*- coding: utf-8 -*-
"""
Created on Fri Feb 26 12:51:27 2016

@author: user11
"""

import os
import glob
import re
import xlwings as xw
import numpy as np
import pandas as pd


color_OK = (208,254,182)
color_NOK = (236,181,167)
path=os.getcwd()
path="G:/01-SUIVI PROCESSUS/GMM/Micro-Vu Vertex/Schrader/43173-820"
fcsv=glob.glob(path+"/CSV Data/*.txt")[0] #File containing CSV
os.system("pause")
fxls=glob.glob(path+"/Layout/*.xlsx")[0] #xlsx file layout for report
os.system("pause")
#print(fcsv,fxls)
#os.system("pause")

def file2df(fname):   
    """CSV import function"""
    
    data = pd.read_csv(fname,header=None,
                 sep=None,
                 index_col=False,)
    return data

wb = xw.Workbook(fxls)
data=file2df(fcsv)
for index, row in data.iterrows():
#    os.system("pause")
    loc=data[0][index]
    value=data[1][index]
    if np.isnan(data[4][index]):
        if np.isnan(data[3][index]):
            color = color_OK
        else:
            color = color_NOK
    elif np.isnan(data[6][index]):
                color = color_OK
    else:
        color = color_NOK       
    #color = (255,255,255)
#    print(loc,value)
    xw.Range(str(loc)).value = value
    xw.Range(str(loc)).color = color
wb.xl_workbook.PrintOut()
#os.system("pause")
wb.close()