# -*- coding: utf-8 -*-
"""
Created on Mon Jul 25 17:02:53 2022

@author: 희토미la
"""

import sys
import chromedriver_autoinstaller
import time as t
from selenium import webdriver
from PyQt5.QtWidgets import QLabel, QPushButton, QWidget, QApplication, QLineEdit, QComboBox
from PyQt5.QtGui import QIcon

class Ui_MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        self.setWindowTitle('헌혈 자가문진 매크로')
        self.setWindowIcon(QIcon('icon.png'))
        # self.resize(300, 360)
        self.setFixedSize(300, 450) # 창 크기 고정
        
        # 매크로 사용 설명서
        self.text_label1 = QLabel(self)
        self.text_label1.move(25, 5)
        self.text_label1.setText("""
해당 프로그램은
헌혈 자가문진을 귀찮아하시는 분들을 위해
제작한 헌혈 자가문진 매크로입니다.
      
※ 주의사항 : 
꼭 헌혈 자가문진의 모든 약간을 
준수하시고 계신 분들만 이용하세요.
                                 """)
        
        # 이름 입력(XXX)
        self.text_label1 = QLabel(self)
        self.text_label1.move(25, 135)
        self.text_label1.setText(' • 이름을 입력해주세요.')

        self.line_edit1 = QLineEdit(self)
        self.line_edit1.move(25,155)
        
        # 주민등록번호 앞자리 입력(123456)
        self.text_label2 = QLabel(self)
        self.text_label2.move(25, 195)
        self.text_label2.setText(' • 주민등록번호 앞자리를 입력해주세요.')

        self.line_edit2 = QLineEdit(self)
        self.line_edit2.move(25,215)
        
        # 주민등록번호 뒷자리 입력(*******)
        self.text_label3 = QLabel(self)
        self.text_label3.move(25, 255)
        self.text_label3.setText(' • 주민등록번호 뒷자리를 입력해주세요.')

        self.line_edit3 = QLineEdit(self)
        self.line_edit3.setEchoMode(QLineEdit.Password)
        self.line_edit3.move(25,275)

        # 지역 선택(수도권,경북,경남,충청,전라,강원/재주) 콤보박
        self.text_label4 = QLabel(self)
        self.text_label4.move(25, 315)
        self.text_label4.setText(' • 지역을 선택해주세요.')
                                 
        self.combobox = QComboBox(self)
        self.combobox.addItem('수도권') # index[0]
        self.combobox.addItem('경북') # index[1]
        self.combobox.addItem('경남') # index[2]
        self.combobox.addItem('충청') # index[3]
        self.combobox.addItem('전라') # index[4]
        self.combobox.addItem('강원/제주') # index[5]
        self.combobox.move(25, 335)
        
        # 확인 버튼
        self.button = QPushButton(self)
        self.button.move(25, 400)
        self.button.setText('확인')
        self.button.clicked.connect(self.button_event)

        self.show()

    def button_event(self):
        name = self.line_edit1.text() # 이름 파라미터 가져오기
        jumin1 = self.line_edit2.text() # 주민등록번호 앞자리 파라미터 가져오기
        jumin2 = self.line_edit3.text() # 주민등록번호 뒷자리 파라미터 가져오기
        combobox = self.combobox.currentIndex() # 콤보박스에서 선택된 항목의 Index 반환
        
        chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]  #크롬드라이버 버전 확인

        try:
            driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe')
        except:
            chromedriver_autoinstaller.install(True)
            driver = webdriver.Chrome(f'./{chrome_ver}/chromedriver.exe')

        driver.implicitly_wait(10)
        
        # 크롬 드라이버로 url(자가문진) 열기
        URL = "https://www.bloodinfo.net/emi2/go_emi4_login_page.do?lang=ko"
        driver.get(URL)

        # 개인 정보 xpath
        name_xpath = '//*[@id="name"]'
        jumin1_xpath = '//*[@id="jumin1"]'
        jumin2_xpath = '//*[@id="jumin2"]'
        button_xpath = '//*[@id="btn_start"]'

        # 개인 정보 자동 입력 및 확인 버튼 클릭
        driver.find_element_by_xpath(name_xpath).send_keys(name)
        driver.find_element_by_xpath(jumin1_xpath).send_keys(jumin1)
        driver.find_element_by_xpath(jumin2_xpath).send_keys(jumin2)
        driver.find_element_by_xpath(button_xpath).click()

        # 자가문진 자동 클릭
        # 확인 클릭
        driver.find_element_by_xpath('//*[@id="btn_next"]').click()
        
        # emi4_document.do
        # 안내문 동의 클릭
        for i in range(2,7):
            driver.find_element_by_xpath('/html/body/div/div[2]/div/div/div[1]/div[1]/div['+ str(i) +']/div/label/span').click()
        # 다음 클릭
        driver.find_element_by_xpath('//*[@id="btn_next2"]').click()
        
        # emi4_quesiton.do
        # 해당없음(다음으로 이동) 클릭
        for i in range(1,12):
            driver.find_element_by_xpath('//*[@id="questno_yn_'+ str(i) +'"]').click()
        # 약간 동의함 클릭
        driver.find_element_by_xpath('//*[@id="area_agreement"]/div/div[1]/div[1]/label/span').click()
        # 제출하기 클릭
        driver.find_element_by_xpath('//*[@id="pageEnd"]/a[2]').click()
        
        # emit4_save.do
        # Q01
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[2]/div[1]/label/span').click()
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[2]/div[2]/label/span').click()
        # Q02
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/table/tbody/tr[1]/td[1]/label/span').click()
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/table/tbody/tr[1]/td[2]/label/span').click()
        # Q03
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[7]/div[2]/label/span').click()
        # Q04
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[8]/div[2]/label/span').click()
        # Q05
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[10]/div[2]/label/span').click()
        # Q06
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[12]/div[1]/label/span').click()
        driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[12]/div[2]/label/span').click()
        # Q07 지역 선택
        #    |   1    |  2   | 3
        # 15 | 수도권 | 경북 | 경남
        # 16 | 충청   | 전라 | 강원/제주
        # ex) 강원/제주 //*[@id="surveyForm"]/div/div[1]/div[16]/div[3]/label/span
        if combobox == 0: # 수도권
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[15]/div[1]/label').click()
        elif combobox == 1: # 경북
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[15]/div[2]/label').click()
        elif combobox == 2: # 경남
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[15]/div[3]/label').click()
        elif combobox == 3: # 충청
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[16]/div[1]/label').click()
        elif combobox == 4: # 전라
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[16]/div[2]/label').click()
        elif combobox == 5: # 강원/제
            driver.find_element_by_xpath('//*[@id="surveyForm"]/div/div[1]/div[16]/div[3]/label').click()
        # 제출하기 클릭
        driver.find_element_by_xpath('//*[@id="btn_submit"]')

if __name__=="__main__":
    app = QApplication(sys.argv)
    ui = Ui_MainWindow()
    
    sys.exit(app.exec_()) # PyQt5 UI 종료 