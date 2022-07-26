# -*- coding: utf-8 -*-
"""
Created on Mon Jul 25 17:02:53 2022

@author: 희토미la, wodns

v1.0.1
"""

import sys, os, time
from threading import Thread
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PyQt5.QtWidgets import QLabel, QPushButton, QWidget, QApplication, QLineEdit, QComboBox, QCheckBox
from PyQt5.QtGui import QIcon, QCloseEvent



def file_path(relative_path, bool=True):
    if getattr(sys, 'frozen', False): # packaging
        if bool:
            APP_EXE_DIR = os.path.dirname(os.path.abspath(sys.executable))
            return os.path.join(APP_EXE_DIR, relative_path)
        else:
            APP_DATA_DIR = sys._MEIPASS
            return os.path.join(APP_DATA_DIR, relative_path)
    else: # script
        APP_EXE_DIR = os.path.dirname(os.path.abspath(__file__))
        APP_DATA_DIR = APP_EXE_DIR
        return os.path.join(APP_EXE_DIR, relative_path)



class Ui_MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        self.setWindowTitle('헌혈 자가문진 매크로')
        self.setWindowIcon(QIcon(file_path('icon.png', False)))
        # self.resize(300, 360)
        self.setFixedSize(300, 475) # 창 크기 고정

        # 매크로 사용 설명서
        self.text_readme = QLabel(self)
        self.text_readme.move(25, 5)
        self.text_readme.setText(
"""
해당 프로그램은
헌혈 자가문진을 귀찮아하시는 분들을 위해
제작한 헌혈 자가문진 매크로입니다.
"""
        )

        self.text_caution = QLabel(self)
        self.text_caution.move(25, 60)
        self.text_caution.setText(
"""
※ 주의사항 : 
꼭 헌혈 자가문진의 모든 약관을 
준수하시고 계신 분들만 이용하세요.
"""
        )
        self.text_caution.setStyleSheet('color: #f2484b; font-weight: bold;')

        # 이름 입력(XXX)
        self.text_name = QLabel(' • 이름을 입력해주세요.', self)
        self.text_name.move(25, 135)

        self.line_name = QLineEdit(self)
        self.line_name.move(25,155)

        # 주민등록번호 앞자리 입력(123456)
        self.text_front = QLabel(' • 주민등록번호 앞자리를 입력해주세요.', self)
        self.text_front.move(25, 195)

        self.line_front = QLineEdit(self)
        self.line_front.move(25,215)

        # 주민등록번호 뒷자리 입력(*******)
        self.text_back = QLabel(' • 주민등록번호 뒷자리를 입력해주세요.', self)
        self.text_back.move(25, 255)

        self.line_back = QLineEdit(self)
        self.line_back.move(25,275)
        self.line_back.setEchoMode(QLineEdit.EchoMode.Password)

        # 주민등록번호 보이기 선택
        self.checkbox = QCheckBox('보이기', self)
        self.checkbox.move(25, 300)
        self.checkbox.setChecked(False)
        self.checkbox.stateChanged.connect(self.checkbox_event)

        # 지역 선택(수도권,경북,경남,충청,전라,강원/재주) 콤보박스
        self.text_region = QLabel(' • 지역을 선택해주세요.', self)
        self.text_region.move(25, 340)

        self.combobox = QComboBox(self)
        self.combobox.move(25, 360)
        self.combobox.addItem('수도권') # index[0]
        self.combobox.addItem('경북') # index[1]
        self.combobox.addItem('경남') # index[2]
        self.combobox.addItem('충청') # index[3]
        self.combobox.addItem('전라') # index[4]
        self.combobox.addItem('강원/제주') # index[5]

        # 확인 버튼
        self.button = QPushButton('확인', self)
        self.button.move(25, 425)
        self.button.resize(150, 25)
        self.button.setAutoDefault(True)
        self.button.clicked.connect(self.button_event)

        # 탭 버튼 순서
        self.setTabOrder(self.line_name, self.line_front)
        self.setTabOrder(self.line_front, self.line_back)
        self.setTabOrder(self.line_back, self.checkbox)
        self.setTabOrder(self.checkbox, self.combobox)
        self.setTabOrder(self.combobox, self.button)
        self.setTabOrder(self.button, self.line_name)

        self.show()

    def checkbox_event(self):
        if self.checkbox.isChecked():
            self.line_back.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.line_back.setEchoMode(QLineEdit.EchoMode.Password)

    def button_event(self):
        croll = Thread(target=self.crolling, daemon=True)
        croll.start()
        self.button.setDisabled(True)

    def crolling(self):
        name = self.line_name.text() # 이름 파라미터 가져오기
        jumin1 = self.line_front.text() # 주민등록번호 앞자리 파라미터 가져오기
        jumin2 = self.line_back.text() # 주민등록번호 뒷자리 파라미터 가져오기
        combobox = self.combobox.currentIndex() # 콤보박스에서 선택된 항목의 Index 반환

        chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]  #크롬드라이버 버전 확인
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        ''' chromedriver console hide 필요 '''
        try:
            self.driver = webdriver.Chrome(service=Service(f'./{chrome_ver}/chromedriver.exe'), options=options)
        except:
            chromedriver_autoinstaller.install(True)
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        self.driver.implicitly_wait(10)

        # 크롬 드라이버로 url(자가문진) 열기
        URL = "https://www.bloodinfo.net/emi2/go_emi4_login_page.do?lang=ko"
        self.driver.get(URL)
        time.sleep(0.1) # 안정화 필요
        try:
            # 개인 정보 자동 입력 및 확인 버튼 클릭
            self.driver.find_element(By.XPATH, '//*[@id="name"]').send_keys(name)
            self.driver.find_element(By.XPATH, '//*[@id="jumin1"]').send_keys(jumin1)
            self.driver.find_element(By.XPATH, '//*[@id="jumin2"]').send_keys(jumin2)
            self.driver.find_element(By.XPATH, '//*[@id="btn_start"]').click()
            
            ''' 문제 확인(뒷 페이지의 확인 안눌림) '''
            # 이미 자가문진을 했다면 종료
            #if self.driver.find_element(By.XPATH, '//*[@id="btn_exit"]'):
            #    self.driver.find_element(By.XPATH, '//*[@id="btn_exit"]').click()
            #    ''' 종료 후 ui 처리 필요 '''
            ''' 팝업에 대한 처리 필요 '''

            # 자가문진 자동 클릭
            # 확인 클릭
            self.driver.find_element(By.XPATH, '//*[@id="btn_next"]').click()

            # emi4_document.do
            # 안내문 동의 클릭
            for i in range(2,7):
                self.driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/div[1]/div[1]/div['+ str(i) +']/div/label/span').click()
            # 다음 클릭
            self.driver.find_element(By.XPATH, '//*[@id="btn_next2"]').click()
            # emi4_quesiton.do
            # 해당없음(다음으로 이동) 클릭
            for i in range(1,12):
                self.driver.find_element(By.XPATH, '//*[@id="questno_yn_'+ str(i) +'"]').click()
            # 약간 동의함 클릭
            self.driver.find_element(By.XPATH, '//*[@id="area_agreement"]/div/div[1]/div[1]/label/span').click()
            # 제출하기 클릭
            self.driver.find_element(By.XPATH, '//*[@id="pageEnd"]/a[2]').click()

            # emit4_save.do
            # Q01
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[2]/div[1]/label/span').click()
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[2]/div[2]/label/span').click()
            # Q02
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/table/tbody/tr[1]/td[1]/label/span').click()
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/table/tbody/tr[1]/td[2]/label/span').click()
            # Q03
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[7]/div[2]/label/span').click()
            # Q04
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[8]/div[2]/label/span').click()
            # Q05
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[10]/div[2]/label/span').click()
            # Q06
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[12]/div[1]/label/span').click()
            self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[12]/div[2]/label/span').click()
            # Q07 지역 선택
            #    |   1    |  2   | 3
            # 15 | 수도권 | 경북 | 경남
            # 16 | 충청   | 전라 | 강원/제주
            # ex) 강원/제주 //*[@id="surveyForm"]/div/div[1]/div[16]/div[3]/label/span
            if combobox == 0: # 수도권
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[15]/div[1]/label').click()
            elif combobox == 1: # 경북
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[15]/div[2]/label').click()
            elif combobox == 2: # 경남
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[15]/div[3]/label').click()
            elif combobox == 3: # 충청
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[16]/div[1]/label').click()
            elif combobox == 4: # 전라
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[16]/div[2]/label').click()
            elif combobox == 5: # 강원/제
                self.driver.find_element(By.XPATH, '//*[@id="surveyForm"]/div/div[1]/div[16]/div[3]/label').click()
            # 제출하기 클릭
            self.driver.find_element(By.XPATH, '//*[@id="btn_submit"]')
        except:
            pass
        finally:
            self.driver.quit()
            self.button.setEnabled(True)

    ''' TypeError 처리 필요 '''
    def closeEvent(self):
        try:
            self.driver.quit()
        except AttributeError:
            pass
        QCloseEvent.accept()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = Ui_MainWindow()
    app.exec() # PyQt5 UI 종료