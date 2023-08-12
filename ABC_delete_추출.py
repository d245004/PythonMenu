import pandas as pd
import os
text = pd.read_excel('C:\\Users\\Jaeri\\Downloads\\song_sale.xls',header=10,usecols='c,o,m,t')
text = text[text['특이사항'].str.contains('ABC',na=False)]
text = text.pivot_table(text,index=['업체명','증표번호','업체(M)'],aggfunc='count')
text.to_excel('C:\\Users\\Jaeri\\Downloads\\ABC_delete.xlsx')
print("작업 완료!")
os.startfile("C:\\Users\\Jaeri\\Downloads\\ABC_delete.xlsx")
