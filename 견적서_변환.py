import tkinter
import sqlalchemy
import oracledb
import pandas as pd
from sqlalchemy import create_engine 
import os
import datetime
import numpy as np  
import time
from tkinter import filedialog
from tkinter import messagebox
from tkinter import * 
import PySimpleGUI as sg
import sys

engine = create_engine('oracle+oracledb://newhaimsweb:newhaims@192.168.18.10:1521/orcl')
connection = oracledb.connect(user="newhaimsweb", password="newhaims",dsn="Autonet_03-PC:1521/orcl")
cursor = connection.cursor()

print('데이터베이스 연결됨')
messagebox.showwarning("경고", "프로그램을 종료합니다")  
sys.exit('종료')

d_path = "c:\\Users\\hanyang\\Downloads\\"
while True:
    #files 변수에 선택 파일 경로 넣기
    file_1 = filedialog.askopenfilename(initialdir=d_path,\
            title = "파일을 선택 해 주세요",\
            filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
    #파일 선택 안했을 때 메세지 출력
    if (len(file_1) == 0):
        messagebox.showwarning("경고", "프로그램을 종료합니다")  
        sys.exit('종료')
    else:
        break



def abc():
    
    sg.theme('DefaultNoMoreNagging')

    layout = [ [sg.Text('헤더가 위치한 행번호에서 -1 빼고 입력 ',font='맑은고딕 10',size=(30,1)),sg.InputText('2',key='hdaa')] ,
               [sg.Text('skip 하기위한 숫자를 입력 default = 0',font='맑은고딕 10',size=(30,1)),sg.InputText('0',key='abc')],
               [sg.Text('취합 할 헤더 타이틀을 입력',font='맑은고딕 10',size=(30,1)),sg.InputText('',key='sbd')],
               [sg.Text('사용 할 열 명칭을 입력',font='맑은고딕 10',size=(30,1)),sg.InputText('partno qty gyun_price',key='plp')],
               [sg.Button('입력')],
               [sg.Button('프로그램 끝내기')]
              ]
    window = sg.Window('옵션 입력', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED: 
            messagebox.showwarning("경고", "프로그램을 종료합니다") 
            sys.exit('프로그램 종료')
            
        if event == '프로그램 끝내기':
            messagebox.showwarning("경고", "프로그램을 종료합니다") 
            sys.exit('프로그램 종료')

        if event == '입력':
            # print(values['hdaa'])
            imsiprint = values['hdaa']
            i_abc = values['abc']
            i_sbd = values['sbd']
            i_plp = values['plp']
            break
        
        print(event)
    window.close()
    return (imsiprint,i_abc,i_sbd,i_plp)

while True:
    lala = abc()
    if lala[2] == '':
        continue
    else:
        break
    
    
int_header = int(lala[0])
int_skiprows = int(lala[1])
int_usecols = str(lala[2]).split()
list_cna = str(lala[3]).split()


# int_header = int(input('header 엑셀 열번호에서 -1 해야됨 __ 숫자 입력='))
# int_skiprows = int(input('skiprows __ 기본 0 으로 입력__숫자='))
# int_usecols = input('usecols 엑셀의 열이름을 입력하시요__ 문자로 입력 =').split()
# # list_cna = ['partno','qty','gyun_price']
# list_cna = input('열이름을 partno qty gyun_price 이렇게 입력하시요!').split()



ddx = pd.read_excel(file_1,header= int_header,skiprows= int_skiprows ,usecols= int_usecols)   # 열 자체를 지정해서 가져오자. 
# cna = ['no','partno','name','qty','gyun_price']   # 열이름 변경 하는 방법
cna = list_cna
ddx.columns = cna
# ddx = ddx.drop(columns=["no","name"])   # 열 삭제하는 방법
now =datetime.datetime.now()
ddx["in_time"] = now
# ddx.to_sql(name='gyun_work',con=engine, if_exists='append',index=False)    # 기존 데이터에 추가하는 옵션
ddx = ddx.to_sql(name='gyun_work',con=engine, if_exists='replace',index=False,dtype={'partno': sqlalchemy.types.VARCHAR(20),'qty':sqlalchemy.types.Integer(),'gyun_price':sqlalchemy.types.Integer(),'in_time':sqlalchemy.DateTime()})     # 기존 테이블을 삭제하고 새로이 추가하는 옵션


print('데이터 추가 완료')
engine.execute("UPDATE gyun_work set partno = replace(partno,'-',' ')")
engine.execute("UPDATE gyun_work set partno = replace(partno,' ','')")
engine.execute("UPDATE gyun_work set partno = replace(partno,'DS','-DS')")
engine.execute("UPDATE gyun_work set partno = replace(partno,'SJ','-SJ')")
print('4가지 조건 업데이트 완료')
engine.execute('DROP TABLE gyun  ')
engine.execute('create table gyun as    '
                'select * from  '
                '(  '
                'select part.lep,   '
                '	  part.partno , '
                '    part.afterptno ,   '
                '    part.abcd ,    '
                '    part.group_cd, '
                '    part.detcd,    '
                '    part.ptnme,    '
                '    part.ptnmh,    '
                '    part.uprice,   '
                '    part.gyun_price,   '
                '    part.vhcknd,   '
                "    DECODE(part.uprice,part.gyun_price,'   ok','   N'),    "
                '    part.invqty,   '
                '    part.qty,  '
                '    part.amdqty1,  '
                '    part.amdqty2,  '
                '    loc.locate1,   '
                '    part.leppartno '
                'from   '
                '  ( (select a.lep,a.partno,b.partno as abcd,a.vhcknd,a.afterptno,a.group_cd,a.detcd,a.ptnme,a.ptnmh,a.uprice,b.gyun_price,b.qty,a.invqty,a.amdqty1,a.amdqty2,a.leppartno   '
                '    from hpm a, gyun_work b    '
                "    where a.partno(+) = b.partno and a.lep(+)<>'1' "
                '   ) ) part,   '
                "	( select locate1,leppartno from hpm_housegrp where lep <>'1') loc   "
                '	where loc.leppartno(+) = part.leppartno '
                '	order by part.abcd  '
                ')  x1 , hpm_gu x2 where x1.detcd = x2.gu(+)  ')
engine.execute('DELETE FROM GYUN    '
                'WHERE ROWID IN (   '
                '      SELECT ROWID FROM (  '
                '             SELECT * FROM (   '
                '                    SELECT ROW_NUMBER() OVER(PARTITION BY leppartno ORDER BY leppartno) AS num '
                '                      FROM gyun    '
                '                    )  '
                '       WHERE num > 1   '
                '      )    '
                '    )  ') 
print('GYUN 지우고 다시 생성함')
engine.execute('DROP TABLE junmun_gyun  ')
engine.execute('CREATE TABLE junmun_gyun AS '
                'SELECT * FROM gyun '
                "WHERE gu <> '   ' AND ROWID IN (SELECT MAX(ROWID) FROM gyun GROUP BY partno)  ")
print('JUNMUN 지우고 생성함')
engine.execute('DROP TABLE jaego_gyun  ')
engine.execute('CREATE TABLE jaego_gyun AS select * from gyun where gu is null  ')
print('JAEGO 지우고 생성함')

# 여러가지 뻘짓을 했다.
# sql문은 실행 하지 못한다(한글이 포함되어있으면)
# 여러개 sql 실행은 코드 짜기 귀찮으니 bat 파일로 실행하자
# os.chdir('C:\\Users\\hanyang')
# os.system("start")
# subprocess.call('C:\\Users\\hanyang\\gyun_file.bat')
# subprocess.run('C:\\Users\\hanyang\\gyun_file.bat',shell=True,encoding='ANSI')
# win32api.WinExec('C:\\Users\\hanyang\\gyun_file.bat')
# sqlplus 와 한글 조합이 맞지않는다. 그냥 오라클 에서 실행 해라. 만지기도 귀찮다. 하나 해결 하면 다른 문제가 또 등장 할지 오르니 !!!!!!!!!!!!
# 해결했다. 배치파일과 sql파일을 ANSI로 저장하면 제대로 실행 된다. 끝이 없네. 이놈의 컴퓨터는. 
# os.startfile('C:\\Users\\hanyang\\gyun_file.bat')
# time.sleep(120)
# sql 작업 실행 한후 작업 할 것

gyun = pd.read_sql_table("gyun",engine)
ss = gyun
two = gyun[gyun.duplicated(['partno','qty'],keep=False) ==True]
ss = ss.drop_duplicates(['partno','qty'],keep=False)
kia = two.query("lep == 'K' & invqty > 0" )
hd = two.query("lep == 'H' & invqty > 0" )
zz = two.query("lep == 'H' & invqty == 0" )
tt = pd.concat([hd,kia,zz])
gyun = pd.concat([ss,tt])
gyun['partno'] = gyun['partno'].str.slice(start=0,stop=5)+'  '+gyun['partno'].str.slice(start=5,stop=15)
gyun['locate1'] = gyun['locate1'].str.slice(start=0 ,stop=3 )+'-'+gyun['locate1'].str.slice(start=3 ,stop=5 )+'-'+\
                    gyun['locate1'].str.slice(start=5 ,stop=7 )+'-'+gyun['locate1'].str.slice(start=7 ,stop=9 )+'-'+\
                    gyun['locate1'].str.slice(start=9 ,stop=11 )

jaego_gyun = pd.read_sql_table("jaego_gyun",engine)
ss1 = jaego_gyun
two1 = jaego_gyun[jaego_gyun.duplicated(['partno','qty'],keep=False) ==True]
ss1 = ss1.drop_duplicates(['partno','qty'],keep=False)
kia1 = two1.query("lep == 'K' & invqty > 0" )
hd1 = two1.query("lep == 'H' & invqty > 0" )
zz1 = two1.query("lep == 'H' & invqty == 0" )
tt1 = pd.concat([hd1,kia1,zz1])
jaego_gyun = pd.concat([ss1,tt1])
jaego_gyun['partno'] = jaego_gyun['partno'].str.slice(start=0,stop=5)+'  '+jaego_gyun['partno'].str.slice(start=5,stop=15)
jaego_gyun['locate1'] = jaego_gyun['locate1'].str.slice(start=0 ,stop=3 )+'-'+jaego_gyun['locate1'].str.slice(start=3 ,stop=5 )+'-'+\
                    jaego_gyun['locate1'].str.slice(start=5 ,stop=7 )+'-'+jaego_gyun['locate1'].str.slice(start=7 ,stop=9 )+'-'+\
                    jaego_gyun['locate1'].str.slice(start=9 ,stop=11 )

junmun_gyun = pd.read_sql_table("junmun_gyun",engine)
junmun_gyun = junmun_gyun.drop_duplicates(['partno','qty'],keep=False)
junmun_gyun['partno'] = junmun_gyun['partno'].str.slice(start=0,stop=5)+'  '+junmun_gyun['partno'].str.slice(start=5,stop=15)
junmun_gyun['locate1'] = junmun_gyun['locate1'].str.slice(start=0 ,stop=3 )+'-'+junmun_gyun['locate1'].str.slice(start=3 ,stop=5 )+'-'+\
                    junmun_gyun['locate1'].str.slice(start=5 ,stop=7 )+'-'+junmun_gyun['locate1'].str.slice(start=7 ,stop=9 )+'-'+\
                    junmun_gyun['locate1'].str.slice(start=9 ,stop=11 )

print('3항목 데이터 가져옴')

gyun.to_excel(d_path + '견적서.xlsx')
jaego_gyun.to_excel(d_path + '견적_재고.xlsx')
junmun_gyun.to_excel(d_path + '견적_전문점.xlsx')
print('3개의 Excel File 생성')
os.startfile(d_path + '견적서.xlsx')
# os.startfile(d_path +'견적_재고.xlsx')
# os.startfile(d_path + '견적_전문점.xlsx')

# 아! 이 지겨운 짓거리도 이제는 끝나나보다. 수고했다.내게.
print(' ---------------  The End  -------------------')