from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import QSettings

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

# 크롬 드라이버 자동 업데이트
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.logger import set_logger
import logging

import pyautogui
import pyperclip
import sys

import os
import traceback
from dotenv import load_dotenv
import time
import csv

# from subprocess import CREATE_NO_WINDOW
# os.getenv()
# logging.log()
#.env 파일 꼭 파이썬과 같은 디렉토리!!

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_PATH = "design4_1.ui"

class MainDialog(QDialog):
    def __init__(self):
        QDialog.__init__(self, None)
        uic.loadUi(os.path.join(BASE_DIR, UI_PATH), self)
        self.keyword:QLineEdit
        self.place_max:QSpinBox
        self.status:QTextBrowser
        self.start_btn:QPushButton
        self.reset_btn:QPushButton
        self.quit_btn:QPushButton
        self.location_btn:QPushButton
        self.remove_btn:QPushButton
        self.cur_location_btn:QPushButton

        #저번에 저장된 디렉토리 설정값 불러오기
        self.settings = QSettings("MY", "crawling_navermap_place")
        self.directory = self.settings.value("directory", os.getcwd())
        self.status.append(f"현재 설정 디렉토리: {self.directory}")
        print(f"현재 설정 디렉토리: {self.directory}")

        self.keyword.setText("대구 맛집")
        self.place_max.setValue(2)


        # 1) 버튼 클릭 이벤트
        # self.객체이름.clicked.connect(self.실행할 함수이름)
        self.start_btn.clicked.connect(self.start)
        self.reset_btn.clicked.connect(self.reset)
        self.quit_btn.clicked.connect(self.quit)
        self.location_btn.clicked.connect(self.selectDirectory)
        self.remove_btn.clicked.connect(self.removeSettingsForTest)
        self.cur_location_btn.clicked.connect(self.getCurrentLocationForTest)
        
        # 로거 생성
        self.logger = logging.getLogger("MY")
        self.logger.setLevel(logging.ERROR)

        formatter = logging.Formatter("%(asctime)s %(levelname)s:%(message)s")
        handler = logging.StreamHandler()
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

        # exe파일일 경우 환경변수 가져오고(이 폴더에 있는 env파일 참고) 파일 로거 설정함
        if getattr(sys, 'frozen', False):
            load_dotenv(dotenv_path=os.path.join(BASE_DIR, '.env'), verbose=True)

            formatter_f = logging.Formatter("%(asctime)s %(levelname)s:%(message)s")
            handler_f = logging.FileHandler("log_event.log")
            handler_f.setFormatter(formatter_f)
            self.logger.addHandler(handler_f)

            set_logger(self.logger) # WDM 라이브러리에 이 로거 보냄
            # print(int(os.getenv("WDM_PROGRESS_BAR")))
            # print(int(os.getenv("GGG", 2)))
            # print("test")

    def getCurrentLocationForTest(self):
        self.status.append(f"현재 디렉토리: {self.directory}")

    def removeSettingsForTest(self):
        self.settings.remove("directory")
        self.directory = os.getcwd()
        self.status.append(f"파일의 저장위치가 초기화됬습니다: {self.directory}")

    def selectDirectory(self):
        selecting_directory = QFileDialog.getExistingDirectory(self, "결과 파일의 저장위치를 설정하세요")
        if selecting_directory: #사용자가 취소버튼이나 x버튼 안누르고 확인버튼 눌렀을 때
            self.directory = selecting_directory.replace("/", "\\", -1)
            self.settings.setValue("directory", self.directory)
            self.status.append(f"파일의 저장위치가 {self.directory}로 설정되었습니다.")
            print(f"변경된 설정값: {self.directory}")
        else:
            self.status.append(f"파일의 저장위치가 변경되지 않았습니다.")

    def reset(self): # 프로그램에 맞게 비울거 추가
        self.keyword.setText("")
        self.place_max.setValue(0)
        # self.keyword.setText("")
        # self.req_comment.setText("")
        # self.last_id_num.setValue(1)
        self.status.setText("리셋이 완료되었습니다.")

    def quit(self):
        sys.exit()

    def start(self):
        # 2) 입력창 텍스트 값 추출
        # self.객체이름.text()
        # input_id = self.id.text()
        # input_pw = self.pw.text()
        input_keyword = self.keyword.text()
        input_place_max = self.place_max.value()

        if input_keyword == "":
            self.status.append("모든 칸을 채워주세요.")
            return 0

        self.status.clear()
        self.status.setText("자동화를 시작합니다.\n자동화가 진행 중인 웹페이지에 마우스와 키보드 입력을 하지 마세요.")
        # self.status.append("로그인 진행중...")
        QApplication.processEvents()

        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)

        # 불필요한 에러 메시지 없애기
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        try:
            # 브라우저 켜기
            path = ChromeDriverManager().install()
            print(path)
            service = Service(executable_path=path)
            # service.creationflags = CREATE_NO_WINDOW // 콘솔 실행없이 크롬 드라이버 서브프로세스를 실행한다?
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except:
            self.logger.error(traceback.format_exc())
            

        # 웹페이지 해당 주소 이동
        driver.implicitly_wait(5) # 웹페이지가 로딩 될때까지 5초는 기다림
        # driver.maximize_window() # 화면 최대화

        driver.get(f"https://map.naver.com/v5/search/{input_keyword}")
        time.sleep(4)

        driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#searchIframe"))

        #크롤링한 맛집 개수 관련 변수 설정
        if input_place_max == 0:
            input_place_max = -1
        count = 0

        #파일 생성
        filename = f"{self.directory}\{input_keyword}_{input_place_max}.csv"
        f = open(filename, 'w', encoding='CP949', newline='')
        csv_writer = csv.writer(f)
        csv_writer.writerow(["이름", "종류", "주소", "전화번호", "평점"])

        #크롤링
        while True:
            #무한 스크롤
            scrollbar = driver.find_element(By.CSS_SELECTOR, "#_pcmap_list_scroll_container")
            before_h = driver.execute_script("return arguments[0].scrollHeight", scrollbar)
            # print(scrollbar.is_enabled())
            # print(scrollbar.is_displayed())
            while True:
                driver.execute_script('arguments[0].scrollBy(0, arguments[0].scrollHeight)', scrollbar)
                # driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollbar)

                time.sleep(1)

                after_h = driver.execute_script("return arguments[0].scrollHeight", scrollbar)

                if before_h == after_h:
                    break
                else:
                    before_h = after_h
            driver.execute_script('arguments[0].scrollBy(0, -arguments[0].scrollHeight*(3/4))', scrollbar)
            time.sleep(2)

            #크롤링
            place_list = driver.find_elements(By.CSS_SELECTOR, "#_pcmap_list_scroll_container > ul > li")
            
            for place in place_list:
                #.eUTV2.Y89AQ
                # #app-root > div > div.XUrfU > div.zRM9F > a.eUTV2.Y89AQ
                # #app-root > div > div.XUrfU > div.zRM9F > a:nth-child(7)
                # #app-root > div > div.XUrfU > div.zRM9F > a:last-child
                # print(place.text)
                #광고가 아닐 경우 로직 수행
                if place.get_attribute("data-laim-exp-id") == "undefinedundefined":
                    # place.find_element(By.XPATH, "./div[1]/a").click() # < 이것도 가능
                    # if(i == 9):
                    #     time.sleep(1)
                    link = place.find_element(By.CSS_SELECTOR, ".place_bluelink")
                    link.location_once_scrolled_into_view
                    link.click()
                    time.sleep(1)

                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#entryIframe"))

                    name = driver.find_element(By.CSS_SELECTOR, "#_title > span.Fc1rA").text
                    kind = driver.find_element(By.CSS_SELECTOR, ".DJJvD").text
                    address = driver.find_element(By.CSS_SELECTOR, ".LDgIH").text
                    try:
                        p_number = driver.find_element(By.CSS_SELECTOR, ".xlx7Q").text
                    except NoSuchElementException:
                        p_number = "None"
                    try:
                        rank = driver.find_element(By.CSS_SELECTOR, ".PXMot.LXIwF > em").text
                    except NoSuchElementException:
                        rank = "None"
                    # try:
                    #     description = driver.find_element(By.CSS_SELECTOR, ".XtBbS.Bhp7c.f7aZ0").text
                    # except:
                    #     description = "None"

                    count += 1
                    print(count, name, kind, address, p_number, rank)
                    csv_writer.writerow([name, kind, address, p_number, rank])
                    if(count == input_place_max):
                        #다 크롤링 했을 시 로직
                        self.save_file(f, driver, filename)
                        return

                    # iframe 밖으로 나오기
                    driver.switch_to.default_content()
                    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "#searchIframe"))

            next_btn = driver.find_element(By.CSS_SELECTOR, "#app-root > div > div.XUrfU > div.zRM9F > a:last-child")
            if next_btn.get_attribute("class") == "eUTV2 Y89AQ":
                # 다 크롤링 했을 시 로직
                self.save_file(f, driver, filename)
                return
            else:
                next_btn.click()
                time.sleep(2)

    def save_file(self, f, driver, filename):
        #다 크롤링 했을 시 로직
        f.close()
        driver.quit()
        self.status.append("자동화가 완료되었습니다.")
        self.status.append(f"{filename}에 저장되었습니다.")

QApplication.setStyle("fusion")
app = QApplication(sys.argv)
main_dialog = MainDialog()
main_dialog.show()

sys.exit(app.exec_())