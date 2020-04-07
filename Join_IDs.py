# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 10:34:10 2019

@author: mape
"""

import h5py
import numpy as np
import pandas

#ID_StatGeb = NA = 9999
#ID_Gemeinde = NA = 9999
from pathlib import Path
path = Path.home() / 'python32' / 'python_dir.txt'
f = open(path, mode='r')
for i in f: path = i
path = Path.joinpath(Path(r'C:'+path),'Dissertation','Agg_Fehler.txt')
f = path.read_text()
f = f.split('\n')

Datenbank = "C:"+f[2]
Group_IDs = "Sonstige"


##--Join to HDF5--##
f = h5py.File(Datenbank,'r+') ##HDF5-File
gr_IDs = f[Group_IDs]
#dset_IDs = gr_IDs["IDs"]
dset_IDs = gr_IDs["Einwohner"]
a = pandas.DataFrame(np.array(dset_IDs))

##--loop-groups--##
for gr in f:
    if "PRZ_MIV" not in gr: continue
#    if "Potenzial" not in gr: continue
    gr_erg = f[gr]
    print gr
    for i in gr_erg:
        if "_100" not in i: continue
#        if "meter" not in i:continue
    
        print i
        dset_erg = gr_erg[i]
        try: text = dset_erg.attrs["Parameter"]
        except: text = "kommt noch"
        b = pandas.DataFrame(np.array(dset_erg))
#        b = b.drop(["ID_500","ID_1000","ID_5000","ID_Gemeinde","ID_StatGeb"], axis=1)
        ##--joining--##
        erg = pandas.merge(b,a,left_on="ID",right_on="ID")
        erg = erg.rename(columns = {'ID_x':'ID'}) ##Ändere den Spaltennamen, für späteren Merge

#        del erg["ID_y"]
#        del erg["Start_ID"]

        ##--savings--##
        x = np.array(np.rec.fromrecords(erg.values))
        x.dtype.names = tuple(erg.dtypes.index.tolist())
        if i in gr_erg.keys(): del gr_erg[i]       
        d = gr_erg.create_dataset(i, data=x)
        d.attrs.create("Parameter",str(text))
        f.flush()


##--end--##
f.flush()
f.close()



