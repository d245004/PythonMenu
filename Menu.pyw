from asyncio import subprocess
# from cgitb import text
from pickle import TRUE
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
import 로또
import 세금계산서 as vat
import First_One
import subprocess


# import 견적서_변환
# import 전문점_조치건_현대_기아_합치기
# import 모비스_6AM_조회하기
# import ABC_delete건_추출

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("menu.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        
        self.setGeometry(10,40,650,450)
        
        self.pushButton.clicked.connect(self.btnFunction)
        self.pushButton_2.clicked.connect(self.btnFunction_2)
        self.pushButton_3.clicked.connect(self.btnFunction_3)
        self.pushButton_4.clicked.connect(self.btnFunction_4)
        self.pushButton_5.clicked.connect(self.btnFunction_5)
        self.pushButton_6.clicked.connect(self.btnFunction_6)
        self.pushButton_7.clicked.connect(self.btnFunction_7)
        self.pushButton_8.clicked.connect(self.btnFunction_8)
        self.pushButton_9.clicked.connect(self.btnFunction_9)
        self.pushButton_10.clicked.connect(self.btnFunction_10)

        self.lineEdit.setPlaceholderText("발행 일자를 입력하시요")
        self.lineEdit.setMaxLength(8)
        self.lineEdit.returnPressed.connect(self.btnFunction_6)
        
        self.calendarWidget.setGridVisible(True)
        # self.centralWidget.clicked[QtCore.QDate].connect(self.btnFunction_6)
        self.calendarWidget.clicked.connect(self.cal_press)
        
        
    def btnFunction(self) :
        # 파이선에서 파일을 읽을때 아래와 같은 오류가 표시된다면,
        # UnicodeDecodeError: 'cp949' codec can't decode byte 0xe2 in position 6987: illegal multibyte sequence
        # 아래와 같이 파일을 여세요.
        # open('파일경로.txt', 'rt', encoding='UTF8')
        # cp949 코덱으로 인코딩 된 파일을 읽어들일때 요런 문제가 생긴다고 하는 군요.
        
        # f = open('first_one.py','rt',encoding='UTF8')
        # exec(f.read())
        # f.close()
        
        # exec(open('first_one.py','rt',encoding='UTF8').read())
        self.hide()
        First_One.first()
        self.show()
        
    def btnFunction_2(self) :
        # exec(open('커머스매출입력확인.py','rt',encoding='UTF8').read())
        subprocess.call('python 커머스매출입력확인.py')
        
    def btnFunction_3(self) :
        # exec(open('로또.py','rt',encoding='UTF8').read())
        로또.로또()
        
    def btnFunction_4(self) :
        # exec(open('월집계표.py','rt',encoding='UTF8').read())
        subprocess.call('python 월집계표.py')
        
    def btnFunction_5(self) :
        # exec(open('일반매출합계추출.py','rt',encoding='UTF8').read())
        subprocess.call('python 일반매출합계추출.py')
        
    def cal_press(self):
        # text = self.calendarWidget.selectedDate()
        # text = str(text)
        # text = text[19:31]
        # 한줄로 요약하면
        text = str(self.calendarWidget.selectedDate())[19:31]
        a=text.replace(",","")
        a=a.replace(")","")
        a=a.split()
        if len(a[1])==1 :
            a[1] = "0"+a[1]
        if len(a[2]) == 1 :
            a[2] = "0"+a[2]
        text = a[0]+a[1]+a[2]
        
        self.lineEdit.setText(text)
        
    def btnFunction_6(self) :
        # exec(open('세금계산서.py','rt',encoding='UTF8').read())
        imsi = self.lineEdit.text()
        if imsi != "":
            vat.vat_work(imsi)
            self.lineEdit.setText("") 
        else:
            print("날자를 Click 하던가 입력하시요!")
        
        
    def btnFunction_7(self) :
        # subprocess.call("견적서_변환.py", shell=True)
        subprocess.call('python 견적서_변환.py')
        # exec(open('견적서_변환.py','rt',encoding='UTF8').read())
        # exec(open('견적서_변환.py').read())
        # print('개뿔 안 되는데 ')
        
    def btnFunction_8(self) :
        # exec(open('전문점_조치건_현대_기아_합치기.py','rt',encoding='UTF8').read())
        subprocess.call('python 전문점_조치건_현대_기아_합치기.py')
        
    def btnFunction_9(self) :
        # exec(open('모비스_6AM_조회하기.py','rt',encoding='UTF8').read())
        subprocess.call('python 모비스_6AM_조회하기.py')
        
    def btnFunction_10(self) :
        # exec(open('ABC_delete_추출.py','rt',encoding='UTF8').read())
        subprocess.call('python ABC_delete_추출.py')
        

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 
    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 
    #프로그램 화면을 보여주는 코드
    myWindow.show()
    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()



    #그럭저럭 성공한 것 같기도 한데 과연???   
    # xxxxxxxxxxxxxxxxxxxxxxx
        # 되긴 된것 같다
        # 되긴 된것 같다

   

    


   

    
   
