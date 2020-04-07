# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 10:34:10 2019

@author: mape
"""

import h5py
import numpy as np
import pandas
from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'Dissertation','Agg_Fehler.txt')
f = path.read_text()
f = f.split('\n')

##--Parameter--##
Datenbank = "C:"+f[2]
Group_IDs = "Sonstige"

##--Join to HDF5--##
f = h5py.File(Datenbank,'r+') ##HDF5-File
gr_IDs = f[Group_IDs]

##--loop-groups--##
for gr in f:
    if "Potenzial" not in gr: continue
    gr_erg = f[gr]
    print gr
    for i in gr_erg:
        if "OEV"  in i:continue
               
            print i
            dset = gr_erg[i]
            try: text = dset.attrs["Parameter"]
            except: text = "kommt noch"
            erg = pandas.DataFrame(np.array(dset))
            for col in erg.columns:
                if "ID" in col:continue
                name = col.replace("__Sum","")
                name = name.replace("Freizeit","FZ")
                name = name.replace("EW_Zensus","EW")
                name = name.replace("Versorgung","EK")
                name = name.replace("__Expo050","05")
                name = name.replace("__Expo020","02")
                name = name.replace("__UH0","UH0")
                erg = erg.rename(columns = {col:name}) ##Ändere den Spaltennamen, für späteren Merge    
            try: erg = erg[["EW15","EW30","EW60","EW02","EW05","EWUH0","AP15","AP30","AP60","AP02","AP05","APUH0",\
                       "EK15","EK30","EK60","EK02","EK05","EKUH0","FZ15","FZ30","FZ60","FZ02","FZ05","FZUH0",\
                       "ID","ID_500","ID_1000","ID_5000","ID_Gemeinde","ID_StatGeb","EW"]]
            except: erg = erg[["EW15","EW30","EW60","EW02","EW05","EWUH0","AP15","AP30","AP60","AP02","AP05","APUH0",\
                       "EK15","EK30","EK60","EK02","EK05","EKUH0","FZ15","FZ30","FZ60","FZ02","FZ05","FZUH0",\
                       "Start_ID"]]
    
            ##--savings--##
            x = np.array(np.rec.fromrecords(erg.values))
            x.dtype.names = tuple(erg.dtypes.index.tolist())
            if i in gr_erg.keys(): del gr_erg[i]       
            d = gr_erg.create_dataset(i, data=x)
            d.attrs.create("Parameter",str(text))
            f.flush()
            
        if "NMIV" in i:
            continue
            
            
            print i
            dset = gr_erg[i]
            try: text = dset.attrs["Parameter"]
            except: text = "kommt noch"
            erg = pandas.DataFrame(np.array(dset))
            
            try: erg = erg[["EW15rad","EW30rad","EW02rad","EW05rad","AP15rad","AP30rad","AP02rad","AP05rad",\
                       "EK15rad","EK30rad","EK02rad","EK05rad","FZ15rad","FZ30rad","FZ02rad","FZ05rad",\
                       "EW15fuss","EW30fuss","EW02fuss","EW05fuss","AP15fuss","AP30fuss","AP02fuss","AP05fuss",\
                       "EK15fuss","EK30fuss","EK02fuss","EK05fuss","FZ15fuss","FZ30fuss","FZ02fuss","FZ05fuss",\
                       "EW5000","EW1000","AP5000","AP1000","EK5000","EK1000","FZ5000","FZ1000",\
                       "ID","ID_500","ID_1000","ID_5000","ID_Gemeinde","ID_StatGeb","EW"]]
            except: erg = erg[["EW15rad","EW30rad","EW02rad","EW05rad","AP15rad","AP30rad","AP02rad","AP05rad",\
                       "EK15rad","EK30rad","EK02rad","EK05rad","FZ15rad","FZ30rad","FZ02rad","FZ05rad",\
                       "EW15fuss","EW30fuss","EW02fuss","EW05fuss","AP15fuss","AP30fuss","AP02fuss","AP05fuss",\
                       "EK15fuss","EK30fuss","EK02fuss","EK05fuss","FZ15fuss","FZ30fuss","FZ02fuss","FZ05fuss",\
                       "EW5000","EW1000","AP5000","AP1000","EK5000","EK1000","FZ5000","FZ1000",\
                       "Start_ID"]]
            
            ##--savings--##
            x = np.array(np.rec.fromrecords(erg.values))
            x.dtype.names = tuple(erg.dtypes.index.tolist())
            if i in gr_erg.keys(): del gr_erg[i]       
            d = gr_erg.create_dataset(i, data=x)
            d.attrs.create("Parameter",str(text))
            
            
##--end--##
f.flush()
f.close()