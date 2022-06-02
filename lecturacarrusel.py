# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 18:39:47 2022

@author: nikel
"""

import xlsxwriter 
import pandas as pd
from pandas import DataFrame

df = pd.read_excel('211221_List.xlsm')
ll=df.values.tolist()
lekua=0
izena=0
brugakoa=0
antolaketa={}
ro=0
co=0
for zerrenda in ll:
    co=0
    for elem in zerrenda:
        if elem=='Position':
            lekua=co
            brugakoa=ro
        if elem=='Sample ID':
            izena=co
        co+=1
    ro+=1
print(lekua)
ro=0
co=0
for zerrenda in ll:
    if ro>=brugakoa:
        aa='" ' + str(zerrenda[izena]) + ' "'
        antolaketa[aa]=zerrenda[lekua]
    ro+=1
        
    