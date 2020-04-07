# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 10:34:10 2019

@author: mape
"""

import h5py
import numpy as np
import pandas
import xlwt
from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'Dissertation','VarK_Kumulation.txt')
f = path.read_text()
f = f.split('\n')

##--Parameter--##
gew = 0
Datenbank = "C:"+f[0]
Resolution = [["ID_500","500"],["ID_Gemeinde","Gem"],["ID_StatGeb","Stat"],["ID_5000","5km"],["ID_1000","1km"]]




if gew == 0:
    Resolution = [["ID_500","500_ungew"],["ID_Gemeinde","Gem_ungew"],["ID_StatGeb","Stat_ungew"],["ID_5000","5km_ungew"],["ID_1000","1km_ungew"]]
    E_Pfad = 'U:'+f[1]
    
else:
    Resolution = [["ID_500","500"],["ID_Gemeinde","Gem"],["ID_StatGeb","Stat"],["ID_5000","5km"],["ID_1000","1km"]]
    E_Pfad = 'U:'+f[2]


def std(x): return np.std(x)

def std_rel(data,w,w_b,g,bezug):
    data.loc[:,"std_sum"] = ((data[w]-data[w_b])**2)*data[g]        
    rel_data = data.groupby(bezug)["std_sum"].agg([('rel_sum', 'sum')]).reset_index()             
    sum_g = data.groupby(bezug)[g].agg([('g_sum', 'sum')]).reset_index()
    data.drop(['std_sum'],axis=1)
    sum_data = pandas.merge(rel_data,sum_g,on=bezug)
    sum_data.loc[:,"std_rel"] = (sum_data["rel_sum"]/sum_data["g_sum"])**0.5    
    sum_data = sum_data.drop(['rel_sum', 'g_sum'],axis=1)
    return sum_data

##--Join to HDF5--##
f = h5py.File(Datenbank,'r+') ##HDF5-File
wb = xlwt.Workbook()

for Bezug in Resolution:
    
    print Bezug[0]
    ws = wb.add_sheet(Bezug[0])
    c = 0
    for groups in f:
        if "Potenzial" not in groups:continue
        ws.write(0, 1+c, groups)
        ws.write(0, 2+c, "VarK")
        ws.write(0, 3+c, "EWVarK")
        ws.write(0, 4+c, "RelVarK")
        ws.write(0, 5+c, "EWRelVarK")
        print groups
    
        gr = f[groups]
   
        r = 1    
        for i in gr:
            if "100" not in i:continue
            if "meter" in i:continue

            name = i.split("_")[1]+"_"+groups.split("_")[1]
        
            dset = gr[i]
            a = pandas.DataFrame(np.array(dset))
            a.loc[:,"n"] = 1
            if Bezug[0] == "ID_StatGeb":a = a[a["ID_StatGeb"]!=9999] ##no missing Stat Areas
            if Bezug[0] == "ID_Gemeinde":a = a[a["ID_Gemeinde"]!=9999] ##no missing Stat Areas
            data_v = pandas.DataFrame(np.array(gr[i.replace("100",Bezug[1])])) ##name Vergleichssystem
            
            for col in list(a.columns):
                if "ID" in col:continue
                if "EW" == col:continue
                if "n" == col: continue
            
            
                ##--Rel--##
                data = data_v.rename(columns={col: "meanBezug", "Start_ID": Bezug[0]})
                z100 = pandas.merge(a,data,on=Bezug[0])

                stdRelVarK = std_rel(z100,col,"meanBezug","n",Bezug[0])
                stdRelVarK = stdRelVarK.rename(columns={"std_rel": "std_RelVarK"})
                
                       
                ##--Calculations--##
                b = a.groupby(Bezug[0])[col].agg([('std', std),('mean', 'mean'), ('n', 'count')]).reset_index()
                b = b.fillna(0)
                EW = a.groupby(Bezug[0])["EW"].agg([('EW', 'sum')]).reset_index()          
                b = pandas.merge(b,EW,on=Bezug[0])
                b = pandas.merge(b,stdRelVarK,on=Bezug[0])
                b = b[b["n"]>0] #no Areas with than one 100-meter cells 
                
                ##VarK##
                b["cv"] = b["std"]/b["mean"]
                b["cv_ew"] = b["cv"]*b["EW"]  
                ws.write(r, 1+c, col)
                ws.write(r, 2+c, round(b["cv"].mean(),3))
                ws.write(r, 3+c, round(b["cv_ew"].sum()/b["EW"].sum(),3))
                
                ##RelVarK##
                b["cv"] = b["std_RelVarK"]/b["mean"]
                b["cv_ew"] = b["cv"]*b["EW"]  
                b = b.replace([np.inf, -np.inf], 0)
                ws.write(r, 4+c, round(b["cv"].mean(),3))
                ws.write(r, 5+c, round(b["cv_ew"].sum()/b["EW"].sum(),3))

                r+=1
            c+=5

##--end--##
f.flush()
f.close()
wb.save(E_Pfad)