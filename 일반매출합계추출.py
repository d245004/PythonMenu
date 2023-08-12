'''
Group A, Group B, Group C, Group D, Group E 만들어 놓고서 작업을 시작 할 것
'''
import pandas as pd
import os

f_name = ['청구서A','청구서B','청구서C','청구서D','청구서E']
aa = pd.DataFrame()
for na in f_name:
    imsi = pd.read_excel("C:\\Users\\Jaeri\\downloads\\"+na+".xls", header=12 )
    # aa = aa.append(imsi,ignore_index=True)
    aa = pd.concat([aa,imsi])

    
aa = aa[['업체코드', '업체명', '총금액']]
aa.업체코드 = (aa.업체코드.str.replace('-00000','')).str.strip()

aa = pd.pivot_table(aa, index=['업체코드', '업체명'], values='총금액', aggfunc='sum')
aa.reset_index
aa.to_excel("C:\\Users\\Jaeri\\Downloads\\일반매출 합계표.xlsx")

os.startfile("C:\\Users\\Jaeri\\Downloads\\일반매출 합계표.xlsx")

