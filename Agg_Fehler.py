# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 10:34:10 2019
@author: mape
"""

import h5py
import numpy as np
import pandas
import xlwt
pandas.set_option('mode.chained_assignment', None)
from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'Dissertation','VarK.txt')
f = path.read_text()
f = f.split('\n')


#################
##--Parameter--##
gew = 0
export_bezug = "ID_500"
runden = 0
meter = 0
##--Parameter--##
#################


if gew == 1: zent = "gew"
else: zent = "ungew"

if runden ==1: rund = "_gerundet"
else: rund = ""

if meter == 1: met = "_meter"
else: met = ""

##--functions--##
def std(x): return np.std(x)

def rel(data,x,y,bezug):
    data.loc[:,"rel"] = data[x]*data[y]
    rel_data = data.groupby(bezug)["rel"].agg([('rel_sum', 'sum')]).reset_index()
    b_data = data.groupby(bezug)[y].agg([('b_sum', 'sum')]).reset_index()
    sum_data = pandas.merge(rel_data,b_data,on=bezug)
    sum_data.loc[:,"mean_rel"] = sum_data["rel_sum"]/sum_data["b_sum"]
    sum_data = sum_data.drop(['b_sum', 'rel_sum'],axis=1)          
    return sum_data

def std_rel(data,w,w_b,g,bezug):
    data.loc[:,"std_sum"] = ((data[w]-data[w_b])**2)*data[g]        
    rel_data = data.groupby(bezug)["std_sum"].agg([('rel_sum', 'sum')]).reset_index()             
    sum_g = data.groupby(bezug)[g].agg([('g_sum', 'sum')]).reset_index()
    data.drop(['std_sum'],axis=1)
    sum_data = pandas.merge(rel_data,sum_g,on=bezug)
    sum_data.loc[:,"std_rel"] = (sum_data["rel_sum"]/sum_data["g_sum"])**0.5    
    sum_data = sum_data.drop(['rel_sum', 'g_sum'],axis=1)
    return sum_data
 

##--Path--##
Datenbank = "C:"+f[2]
f = h5py.File(Datenbank,'r+') ##HDF5-File
wb = xlwt.Workbook()
modi = ["MIV","Fuss","Rad","OEV"]
if gew == 0:
    Resolution = [["ID_500","500_ungew"],["ID_Gemeinde","Gem_ungew"],["ID_StatGeb","Stat_ungew"],["ID_5000","5km_ungew"],["ID_1000","1km_ungew"]]
    E_Pfad = 'U:'+f[0]+met+rund+'.xls'
    
else:
    Resolution = [["ID_500","500"],["ID_Gemeinde","Gem"],["ID_StatGeb","Stat"],["ID_5000","5km"],["ID_1000","1km"]]
    E_Pfad = 'U:'+f[1]+met+rund+'.xls'

##--Calculation--##
for Bezug in Resolution:
    
    print(Bezug[0])
    ws = wb.add_sheet(Bezug[0])
    ws.write(0, 0, "Anzahl")
    ws.write(0, 1, "Name")        
    ws.write(0, 2, "corr_EW_n_OEV")
    ws.write(0, 3, "corr_VarK_n_OEV")  
    ws.write(0, 4, "corr_VarK_EW_OEV")  
    ws.write(0, 5, "corr_EW_n_MIV")
    ws.write(0, 6, "corr_VarK_n_MIV")  
    ws.write(0, 7, "corr_VarK_EW_MIV")  
    ws.write(0, 8, "corr_EW_n_Fuss")
    ws.write(0, 9, "corr_VarK_n_Fuss")  
    ws.write(0, 10, "corr_VarK_EW_Fuss")  
    ws.write(0, 11, "corr_EW_n_Rad")
    ws.write(0, 12, "corr_VarK_n_Rad")  
    ws.write(0, 13, "corr_VarK_EW_Rad")
    
    ws.write(0, 15, "Name")     
    ws.write(0, 16, "vergl_OEV")
    ws.write(0, 17, "vergl_ew_OEV")
    ws.write(0, 18, "vergl_MIV")
    ws.write(0, 19, "vergl_ew_MIV")
    ws.write(0, 20, "vergl_Fuss")
    ws.write(0, 21, "vergl_ew_Fuss")
    ws.write(0, 22, "vergl_Rad")
    ws.write(0, 23, "vergl_ew_Rad")
    
    ws.write(0, 24, "Name")
    ws.write(0, 25, "ID_OEV")
    ws.write(0, 26, "ID_MIV")
    ws.write(0, 27, "ID_Fuss")
    ws.write(0, 28, "ID_Rad")
    
    ws.write(0, 30, "Name")
    ws.write(0, 31, "Schnitt_OEV")
    ws.write(0, 32, "Schnitt_MIV")
    ws.write(0, 33, "Schnitt_Fuss")
    ws.write(0, 34, "Schnitt_Rad")

    
            
    for groups in f:
        if "Ergebnisse" not in groups:continue
        gr = f[groups]
        r = 1
        print groups
        for i in gr:
            if "100" not in i:continue
            if "meter" in i:continue
            if "PuR_" in i or "Hst_" in i:continue        
            a100 = pandas.DataFrame(np.array(gr[i]))
            a100.loc[:,"n"] = 1
            data_v = pandas.DataFrame(np.array(gr[i.replace("100",Bezug[1])])) ##name Vergleichssystem
            if "_MIV" in groups: ws.write(r, 0, len(np.unique(a100["Ziel_ID"]))) #Only in the first column                            
            if Bezug[0] == "ID_StatGeb":a100 = a100[a100["ID_StatGeb"]!=9999] ##no missing Stat Areas
            if Bezug[0] == "ID_Gemeinde":a100 = a100[a100["ID_Gemeinde"]!=9999] ##no missing Stat Areas            
            
            if "_OEV" in i:
                if meter == 1:continue
                w = "Reisezeit"
                z100 = a100[a100["Reisezeit"]<=60]
                ws.write(r, 31, round(z100.Reisezeit.mean(),3))
                z100 = z100.drop(['StartHst','ZielHst','Anbindungszeit','Abgangszeit','BH'],axis=1)  
                data_v = data_v.drop(['StartHst','ZielHst','Anbindungszeit','Abgangszeit','BH'],axis=1)                
                z100 = z100.rename(columns={"Ziel_ID": "Ziel_100"})
                data_v = data_v.rename(columns={"Ziel_ID": "Ziel_V"})                 
                data_v = data_v.rename(columns={w: "meanBezug", "Start_ID": Bezug[0], "UH": "UHBezug"})                
                z100["RZ_ew"] = z100[w]*z100["EW"]
                m100 = z100.groupby(Bezug[0])[w].agg([('mean100', 'mean'),('n100', 'count')]).reset_index()   
                m100 = m100.fillna(0)   
                EW = z100.groupby(Bezug[0])["EW"].agg([('EW', 'sum')]).reset_index()
                EW = EW[EW["EW"]>0] 
                m100_ew = rel(z100,w,"EW",Bezug[0])
                m100_ew = m100_ew.rename(columns={"mean_rel": "mean100_ew"})
                m100_ew = m100_ew.fillna(0)
                
                ##--merge to R100--##
                z100 = pandas.merge(z100,m100,on=Bezug[0])
                z100 = pandas.merge(z100,m100_ew,on=Bezug[0])
                z100 = pandas.merge(z100,data_v,on=Bezug[0]) ##join accessibility values to 100-meter-resolution
                
                ##--calculate std--##
                std100 = z100.groupby(Bezug[0])[w].agg([('std', std)]).reset_index()
                stdeVarK = std_rel(z100,w,"mean100","EW",Bezug[0])
                stdeVarK = stdeVarK.rename(columns={"std_rel": "std_eVarK"})
                stdRelVarK = std_rel(z100,w,"meanBezug","n",Bezug[0])
                stdRelVarK = stdRelVarK.rename(columns={"std_rel": "std_RelVarK"})
                stdeRelVarK = std_rel(z100,w,"meanBezug","EW",Bezug[0])
                stdeRelVarK = stdeRelVarK.rename(columns={"std_rel": "std_eRelVarK"})
                
                ##--merge to bezug--##
                zbezug = pandas.merge(EW,m100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,m100_ew,on=Bezug[0])
                zbezug = pandas.merge(zbezug,data_v,on=Bezug[0])
                zbezug = pandas.merge(zbezug,std100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdRelVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeRelVarK,on=Bezug[0])
                zbezug = zbezug.replace([np.inf, -np.inf], 0)  
                zbezug = zbezug.fillna(0) 
                
                
                ##--save results--##
                #VarK / EWVarK
                zbezug["cv"] = zbezug["std"]/zbezug["mean100"]
                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
                zbezug = zbezug.replace([np.inf, -np.inf], 0)
                ws.write(r, 1, i.split("_")[1])     
                ws.write(r, 15, i.split("_")[1])
                ws.write(r, 24, i.split("_")[1])  
                ws.write(r, 30, i.split("_")[1]) 
                ws.write(r, 2, round(zbezug['n100'].corr(zbezug['EW']),3))
                ws.write(r, 3, round(zbezug['n100'].corr(zbezug['cv']),3))
                ws.write(r, 4, round(zbezug['EW'].corr(zbezug['cv']),3))
                
                ws.write(r, 16, round((zbezug["mean100"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ws.write(r, 17, round((zbezug["mean100_ew"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ID_Ziel = float((z100["Ziel_100"]!=z100["Ziel_V"]).sum())/len(z100)
                ws.write(r, 25, round(ID_Ziel,3))
                
#                #eVarK / EWeVarK
#                zbezug["cv"] = zbezug["std_eVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 37, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 38, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#
#                #RelVarK / RelVarK
#                zbezug["cv"] = zbezug["std_RelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 39, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 40, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK / EWeRelVarK
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 41, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 42, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK_ew / EWeRelVarK_ew
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100_ew"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 43, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 44, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
                                
    
            if "_MIV" in i:
                w = "tAkt"
                if meter == 1: w = "Meter"
                z100 = a100[a100["tAkt"]<=60]
                ws.write(r, 32, round(z100.tAkt.mean(),3))
                
                z100 = z100.rename(columns={"Ziel_ID": "Ziel_100"})
                data_v = data_v.rename(columns={"Ziel_ID": "Ziel_V"})
#                z100 = z100.drop(['Ziel_ID'],axis=1)  
#                data_v = data_v.drop(['Ziel_ID'],axis=1)
                data_v = data_v.rename(columns={w: "meanBezug", "Start_ID": Bezug[0]})
                if runden == 1:
                    z100 = z100.round(0)
                    data_v = data_v.round(0)                
                z100["RZ_ew"] = z100[w]*z100["EW"]
                m100 = z100.groupby(Bezug[0])[w].agg([('mean100', 'mean'),('n100', 'count')]).reset_index()   
                m100 = m100.fillna(0)
                EW = z100.groupby(Bezug[0])["EW"].agg([('EW', 'sum')]).reset_index()
                EW = EW[EW["EW"]>0] 
                m100_ew = rel(z100,w,"EW",Bezug[0])
                m100_ew = m100_ew.rename(columns={"mean_rel": "mean100_ew"})
                m100_ew = m100_ew.fillna(0)
                
                ##--merge to R100--##
                z100 = pandas.merge(z100,m100,on=Bezug[0])
                z100 = pandas.merge(z100,m100_ew,on=Bezug[0])
                z100 = pandas.merge(z100,data_v,on=Bezug[0]) ##join accessibility values to 100-meter-resolution
                
                ##--calculate std--##
                std100 = z100.groupby(Bezug[0])[w].agg([('std', std)]).reset_index()
                stdeVarK = std_rel(z100,w,"mean100","EW",Bezug[0])
                stdeVarK = stdeVarK.rename(columns={"std_rel": "std_eVarK"})
                stdRelVarK = std_rel(z100,w,"meanBezug","n",Bezug[0])
                stdRelVarK = stdRelVarK.rename(columns={"std_rel": "std_RelVarK"})
                stdeRelVarK = std_rel(z100,w,"meanBezug","EW",Bezug[0])
                stdeRelVarK = stdeRelVarK.rename(columns={"std_rel": "std_eRelVarK"})
                
                ##--merge to bezug--##
                zbezug = pandas.merge(EW,m100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,m100_ew,on=Bezug[0])
                zbezug = pandas.merge(zbezug,data_v,on=Bezug[0])
                zbezug = pandas.merge(zbezug,std100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdRelVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeRelVarK,on=Bezug[0])
                zbezug = zbezug.replace([np.inf, -np.inf], 0)  
                zbezug = zbezug.fillna(0) 
                
                
                ##--save results--##
                #VarK / EWVarK
                zbezug["cv"] = zbezug["std"]/zbezug["mean100"]
                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
                zbezug = zbezug.replace([np.inf, -np.inf], 0)
                
                ws.write(r, 5, round(zbezug['n100'].corr(zbezug['EW']),3))
                ws.write(r, 6, round(zbezug['n100'].corr(zbezug['cv']),3))
                ws.write(r, 7, round(zbezug['EW'].corr(zbezug['cv']),3))
                
                ws.write(r, 18, round((zbezug["mean100"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ws.write(r, 19, round((zbezug["mean100_ew"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ID_Ziel = float((z100["Ziel_100"]!=z100["Ziel_V"]).sum())/len(z100)
                ws.write(r, 26, round(ID_Ziel,3))
                
                
#                #eVarK / EWeVarK
#                zbezug["cv"] = zbezug["std_eVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 4, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 5, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#
#                #RelVarK / RelVarK
#                zbezug["cv"] = zbezug["std_RelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 6, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 7, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK / EWeRelVarK
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 8, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 9, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK_ew / EWeRelVarK_ew
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100_ew"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 10, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 11, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
                
           
            if "_NMIV" in i: ##tFuss
                w = "tFuss"
                if meter == 1: w = "Meter"
                z100 = a100[a100["tFuss"]<=60]
                ws.write(r, 33, round(z100.tFuss.mean(),3))
                z100 = z100.drop(['tRad'],axis=1)  
                z100 = z100.rename(columns={"Ziel_ID": "Ziel_100"})
                data_v = data_v.rename(columns={"Ziel_ID": "Ziel_V"})
                data_v = data_v.rename(columns={w: "meanBezug", "Start_ID": Bezug[0]})
                if runden == 1:
                    z100 = z100.round(0)
                    data_v = data_v.round(0)
                z100["RZ_ew"] = z100[w]*z100["EW"]
                m100 = z100.groupby(Bezug[0])[w].agg([('mean100', 'mean'),('n100', 'count')]).reset_index()   
                m100 = m100.fillna(0)   
                EW = z100.groupby(Bezug[0])["EW"].agg([('EW', 'sum')]).reset_index()
                EW = EW[EW["EW"]>0] 
                m100_ew = rel(z100,w,"EW",Bezug[0])
                m100_ew = m100_ew.rename(columns={"mean_rel": "mean100_ew"})
                m100_ew = m100_ew.fillna(0)
                
                ##--merge to R100--##
                z100 = pandas.merge(z100,m100,on=Bezug[0])
                z100 = pandas.merge(z100,m100_ew,on=Bezug[0])
                z100 = pandas.merge(z100,data_v,on=Bezug[0]) ##join accessibility values to 100-meter-resolution

                ##--calculate std--##
                std100 = z100.groupby(Bezug[0])[w].agg([('std', std)]).reset_index()
                stdeVarK = std_rel(z100,w,"mean100","EW",Bezug[0])
                stdeVarK = stdeVarK.rename(columns={"std_rel": "std_eVarK"})
                stdRelVarK = std_rel(z100,w,"meanBezug","n",Bezug[0])
                stdRelVarK = stdRelVarK.rename(columns={"std_rel": "std_RelVarK"})
                stdeRelVarK = std_rel(z100,w,"meanBezug","EW",Bezug[0])
                stdeRelVarK = stdeRelVarK.rename(columns={"std_rel": "std_eRelVarK"})
                
                ##--merge to bezug--##
                zbezug = pandas.merge(EW,m100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,m100_ew,on=Bezug[0])
                zbezug = pandas.merge(zbezug,data_v,on=Bezug[0])
                zbezug = pandas.merge(zbezug,std100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdRelVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeRelVarK,on=Bezug[0])
                zbezug = zbezug.replace([np.inf, -np.inf], 0)  
                zbezug = zbezug.fillna(0) 
                
                ##--save results--##
                #VarK / EWVarK
                zbezug["cv"] = zbezug["std"]/zbezug["mean100"]
                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
                zbezug = zbezug.replace([np.inf, -np.inf], 0)
                
                ws.write(r, 8, round(zbezug['n100'].corr(zbezug['EW']),3))
                ws.write(r, 9, round(zbezug['n100'].corr(zbezug['cv']),3))
                ws.write(r, 10, round(zbezug['EW'].corr(zbezug['cv']),3))
                
                ws.write(r, 20, round((zbezug["mean100"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ws.write(r, 21, round((zbezug["mean100_ew"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ID_Ziel = float((z100["Ziel_100"]!=z100["Ziel_V"]).sum())/len(z100)
                ws.write(r, 27, round(ID_Ziel,3))
                
#                #eVarK / EWeVarK
#                zbezug["cv"] = zbezug["std_eVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 15, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 16, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#
#                #RelVarK / RelVarK
#                zbezug["cv"] = zbezug["std_RelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 17, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 18, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK / EWeRelVarK
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 19, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 20, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK_ew / EWeRelVarK_ew
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100_ew"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 21, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 22, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                                                              
                
            if "_NMIV" in i: ##tRad   
                w = "tRad"
                if meter == 1: w = "Meter"
                z100 = a100[a100["tRad"]<=60]
                ws.write(r, 34, round(z100.tRad.mean(),3))
                z100 = z100.drop(['tFuss'],axis=1)
                z100 = z100.rename(columns={"Ziel_ID": "Ziel_100"})
                data_v = data_v.rename(columns={"Ziel_ID": "Ziel_V"})

                if meter == 0:
                    data_v = data_v.drop(['meanBezug'],axis=1)
                    data_v = data_v.rename(columns={w: "meanBezug"})  
                if runden == 1:
                    z100 = z100.round(0)
                    data_v = data_v.round(0)
                z100["RZ_ew"] = z100[w]*z100["EW"]
                m100 = z100.groupby(Bezug[0])[w].agg([('mean100', 'mean'),('n100', 'count')]).reset_index()   
                m100 = m100.fillna(0)   
                EW = z100.groupby(Bezug[0])["EW"].agg([('EW', 'sum')]).reset_index()
                EW = EW[EW["EW"]>0] 
                m100_ew = rel(z100,w,"EW",Bezug[0])
                m100_ew = m100_ew.rename(columns={"mean_rel": "mean100_ew"})
                m100_ew = m100_ew.fillna(0)
                
                ##--merge to R100--##
                z100 = pandas.merge(z100,m100,on=Bezug[0])
                z100 = pandas.merge(z100,m100_ew,on=Bezug[0])
                z100 = pandas.merge(z100,data_v,on=Bezug[0]) ##join accessibility values to 100-meter-resolution

                ##--calculate std--##
                std100 = z100.groupby(Bezug[0])[w].agg([('std', std)]).reset_index()
                stdeVarK = std_rel(z100,w,"mean100","EW",Bezug[0])
                stdeVarK = stdeVarK.rename(columns={"std_rel": "std_eVarK"})
                stdRelVarK = std_rel(z100,w,"meanBezug","n",Bezug[0])
                stdRelVarK = stdRelVarK.rename(columns={"std_rel": "std_RelVarK"})
                stdeRelVarK = std_rel(z100,w,"meanBezug","EW",Bezug[0])
                stdeRelVarK = stdeRelVarK.rename(columns={"std_rel": "std_eRelVarK"})
                
                ##--merge to bezug--##
                zbezug = pandas.merge(EW,m100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,m100_ew,on=Bezug[0])
                zbezug = pandas.merge(zbezug,data_v,on=Bezug[0])
                zbezug = pandas.merge(zbezug,std100,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdRelVarK,on=Bezug[0])
                zbezug = pandas.merge(zbezug,stdeRelVarK,on=Bezug[0])
                zbezug = zbezug.replace([np.inf, -np.inf], 0)  
                zbezug = zbezug.fillna(0)           
                                
                ##--save results--##
                #VarK / EWVarK
                zbezug["cv"] = zbezug["std"]/zbezug["mean100"]
                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]
                zbezug = zbezug.replace([np.inf, -np.inf], 0)
                
                ws.write(r, 11, round(zbezug['n100'].corr(zbezug['EW']),3))
                ws.write(r, 12, round(zbezug['n100'].corr(zbezug['cv']),3))
                ws.write(r, 13, round(zbezug['EW'].corr(zbezug['cv']),3))
                
                ws.write(r, 22, round((zbezug["mean100"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ws.write(r, 23, round((zbezug["mean100_ew"]<zbezug["meanBezug"]).sum()/float(len(zbezug)),3))
                ID_Ziel = float((z100["Ziel_100"]!=z100["Ziel_V"]).sum())/len(z100)
                ws.write(r, 28, round(ID_Ziel,3))
                
#                #eVarK / EWeVarK
#                zbezug["cv"] = zbezug["std_eVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 26, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 27, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#
#                #RelVarK / RelVarK
#                zbezug["cv"] = zbezug["std_RelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 28, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 29, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK / EWeRelVarK
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 30, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 31, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
#                
#                #eRelVarK_ew / EWeRelVarK_ew
#                zbezug["cv"] = zbezug["std_eRelVarK"]/zbezug["mean100_ew"]
#                zbezug["cv_ew"] = zbezug["cv"]*zbezug["EW"]                
#                zbezug = zbezug.replace([np.inf, -np.inf], 0)
#                ws.write(r, 32, round(zbezug["cv"].mean(),3))                                               
#                ws.write(r, 33, round(zbezug["cv_ew"].sum()/zbezug["EW"].sum(),3))
                

                         
            r+=1
        
##--end--##
f.flush()
f.close()
wb.save(E_Pfad)