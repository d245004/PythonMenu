import pandas as pd
import numpy as np
import math
import os


path = "C:\\Users\\Jaeri\\Downloads\\"
path_2 = "C:\\Users\\Jaeri\\OneDrive - 오토넷\\문서\\excel Data\\"

hpm_gu = pd.read_excel(path_2+'HPM_GU.XLS')
hpm_gu.rename(columns= {'GU':'품목'},inplace=True)
hd = pd.read_excel(path+"현대_6am.xls",usecols='a,c,d,g,i,m,w,x')
kia = pd.read_excel(path+"기아_6am.xls",usecols='a,c,d,g,i,m,w,x')
hap = pd.concat([hd,kia])

hap = hap[hap['총재고'] < hap['6AMS'] * 0.5]
hap['부족수량'] = (hap['6AMS'] * 0.5) - hap['총재고']
hap['부족수량'] = (np.ceil(hap['부족수량'])).astype(int)
junmun = hap[hap.품목.isin(hpm_gu.품목)]
mobis = hap[~hap.품목.isin(hpm_gu.품목)]
hap.to_excel(path+"test.xlsx")
junmun.to_excel(path+"6AM 전문점 주문할 건 .xlsx")
mobis.to_excel(path+"6AM MOBIS 주문할 건.xlsx")
os.startfile(path+'test.xlsx')
