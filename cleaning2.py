# -*- coding: utf-8 -*-
"""
Spyder Editor"""

#Importing packages for analysis
import pandas as pd
from pathlib import Path


#Key inputs
start_date = '1/4/1999'
end_date = '8/31/2020'

#Uploading file in Pandas

poil = Path('..', 'Marquez Joint', 'OIL_and_ER_vCLEAN.xls') 
poil = pd.read_excel(poil,'Daily')
poil = poil.rename(columns={'DATE':'date'})
vix_cls = Path('..', 'Marquez Joint', 'OIL_and_ER_vCLEAN.xls') 
vix_cls = pd.read_excel(vix_cls,'Daily_Close')
vix_cls = vix_cls.rename(columns={'DATE':'date'})

#Merging poil and vix_cls

poil_clean_a = poil.merge(vix_cls,on='date',how='outer',validate='m:1')


#Consolidating df to start and end date

after_start_date = poil_clean_a['date'] >= start_date
before_end_date = poil_clean_a['date'] <= end_date
between_two_dates = after_start_date & before_end_date
poil_clean_b = poil_clean_a.loc[between_two_dates]


#Taking out columns that do not have 2020 data.  Will add them back later. 
trade_w = poil_clean_b[['date','DTWEXB','DTWEXM','DTWEXO']].copy()
poil_clean_b = poil_clean_b.drop(['DTWEXB','DTWEXM','DTWEXO'],axis=1)
poil_clean_c = poil_clean_b.fillna(method='ffill',limit=5)
trade_w2 = trade_w.fillna(method='ffill',limit=5)
#Flip exchange rates so everything is per USD

poil_clean_c['DEXALUS'] = 1/poil_clean_c['DEXUSAL']
poil_clean_c['DEXEUUS'] = 1/poil_clean_c['DEXUSEU']
poil_clean_c['DEXNZUS'] = 1/poil_clean_c['DEXUSNZ']
poil_clean_c['DEXUKUS'] = 1/poil_clean_c['DEXUSUK']

poil_clean_d = poil_clean_c.merge(trade_w2,on='date',how='outer',validate='m:1')
poil_clean_e = poil_clean_d.drop(['DEXUSAL',	'DEXUSEU','DEXUSNZ',	'DEXUSUK'],axis=1)


poil_clean_e.to_excel(r'/Users/kubij/My Documents/Marquez Joint/test3.xlsx',index=False)

#May use this line later for analysis.  #