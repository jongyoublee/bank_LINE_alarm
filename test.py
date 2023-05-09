import os
from io import BytesIO
# from time import sleep
from time import *
from urllib.request import urlretrieve as download
import pyperclip
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By

import win32.win32clipboard
import pandas as pd
#import win32clipboard  # !pip install pywin32
import win32com.client as win32
from PIL import Image  # !pip install Pillow
from openpyxl import Workbook  # !pip install openpyxl
from selenium import webdriver  # !pip install selenium
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
import pyautogui
from selenium.webdriver.support.ui import Select
#from home_pass import password
#from home_pass import studiopass
import shutil

def execute_script(script): # 자바스크립트 로딩 에러를 방지하기 위한 헬퍼함수
    while True:
        try:
            driver.execute_script(script)
            break
        except JavascriptException:
            sleep(0.5)

def img_doubleclick(img_icon):
    img_filename=img_icon+'.png'
    while True:
        img_icon=pyautogui.locateOnScreen(img_filename)
        if img_icon==None:
            sleep(0.5)
        else:
            img_icon_center = pyautogui.center(img_icon)
            # pyautogui.click(img_icon_center)
            pyautogui.doubleClick(img_icon_center)
            break

def img_click(img_icon):
    img_filename=img_icon+'.png'
    while True:
        img_icon=pyautogui.locateOnScreen(img_filename)
        if img_icon==None:
            sleep(0.5)
        else:
            img_icon_center = pyautogui.center(img_icon)
            # pyautogui.click(img_icon_center)
            pyautogui.click(img_icon_center)
            break

def img_mouse_scroll(img_icon):
    img_filename = img_icon + '.png'
    while True:
        img_icon = pyautogui.locateOnScreen(img_filename)
        if img_icon == None:
            pyautogui.scroll(-5)
        else:
            img_icon_center = pyautogui.center(img_icon)
            # pyautogui.click(img_icon_center)
            pyautogui.doubleClick(img_icon_center)
            break



def img_arrow_click(img_icon):
    img_filename = img_icon + '.png'
    while True:
        img_icon = pyautogui.locateOnScreen(img_filename)
        if img_icon == None:
            img_click('hometax05arrow')
            pyautogui.move(10, None)
            # pyautogui.press('down')
        else:
            img_icon_center = pyautogui.center(img_icon)
            # pyautogui.click(img_icon_center)
            pyautogui.doubleClick(img_icon_center)
            break

def img_arrow_down(img_icon):
    img_filename = img_icon + '.png'
    while True:
        img_icon = pyautogui.locateOnScreen(img_filename)
        if img_icon == None:
            # img_click('hometax05arrow')
            # pyautogui.move(10, None)
            pyautogui.press('down')
        else:
            img_icon_center = pyautogui.center(img_icon)
            # pyautogui.click(img_icon_center)
            pyautogui.doubleClick(img_icon_center)
            break

def match_arrow_down(title):
    while True:
        if driver.find_element_by_css_selector("[title^='"+title+"']") == None:
            # img_click('hometax05arrow')
            # pyautogui.move(10, None)
            print('None')
            pyautogui.press('down')
        else:
            print('else')
            driver.find_element_by_css_selector("[title^='"+title+"']").click
            break

# method to get the downloaded file name
options=webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs={"profile.default_content_settings.popups": 0,
       # "download.default_directory": r"/Users/miero/Downloads",
       "download.default_directory": r"c:/Users/bm115\Downloads",
       # "download.default_directory": r"c:\Users\miero\Downloads\hometax",
       "download.prompt_for_download": False,
       "safebrowsing.enabled": True,
       "directory_upgrade":True}
options.add_experimental_option("prefs",prefs)
#driver=webdriver.Chrome(chrome_options=options)
driver=webdriver.Chrome(executable_path='C:/Users/bm115/Downloads/temp/chromedriver.exe', chrome_options=options)

def getDownLoadedFileName(waitTime):
    driver.execute_script("window.open()")
    # switch to new tab
    driver.switch_to.window(driver.window_handles[-1])
    # navigate to chrome downloads
    driver.get('chrome://downloads')
    # driver.get('chrome://downloads/')
    # define the endTime
    endTime = time()+waitTime
    while True:
        try:
            # get downloaded percentage
            downloadPercentage = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
            # check if downloadPercentage is 100 (otherwise the script will keep waiting)
            if downloadPercentage == 100:
                # return the file name once the download is completed
                return driver.execute_script("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
        except:
            pass
        sleep(1)
        if time() > endTime:
            break
# DOWNLOAD_DIR = 'c:/users/miero/desktop/imgs'

# mill, hempel, studio 선택
target='bumin'

# driver = webdriver.Chrome(r"C:\Users\miero\chromedriver.exe")
# driver = webdriver.Chrome()
driver.implicitly_wait(7)
driver.get('https://www.hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index.xml')
sleep(7)

# 로그인
#driver.find_element_by_css_selector('#group88615548').click()
driver.find_element(By.CSS_SELECTOR, '#group88615548').click()
sleep(3)

# iframe 이동
# iframe=driver.find_element_by_css_selector("html>body>div#wrap>div#container_wrap>iframe#txppIframe")
iframe=driver.find_element(By.CSS_SELECTOR, "#txppIframe")
driver.switch_to.frame(iframe)

# 요소 있는지 검사
# a=driver.find_elements_by_css_selector("input#trigger38.w2trigger")
# print(len(a))

sleep(9)

# 아이디 로그인
# driver.find_element_by_css_selector("input#trigger38.w2trigger ").click()
driver.find_element(By.CSS_SELECTOR, "#group91882156").click()
sleep(5)


# iframe 이동
# iframe=driver.find_element_by_css_selector("iframe#dscert")
iframe=driver.find_element(By.CSS_SELECTOR, "#dscert")
driver.switch_to.frame(iframe)
sleep(9)
# driver.find_element_by_css_selector("a#stg_hdd").click()
driver.find_element(By.CSS_SELECTOR, "#stg_hdd").click()
sleep(1)
# driver.find_element_by_css_selector("li#hdd_driver_2").click()
driver.find_element(By.XPATH, "//a[text()='이동식 디스크 (G)']").click()
sleep(1)
#######################################################
# driver.find_element_by_xpath('//*[@title="주식회사밀앤아이_0000413351"]').click()
# driver.find_element_by_xpath('//*[@title="헴펠_0000308533"]').click()
# driver.find_element_by_css_selector('tr#row0dataTable').click()

# for elem in driver.find_elements_by_xpath('.//span[@title = "주식회사밀앤아이_0000413351"]'):
#     print(elem.text)

if target=='bumin':
    target_title='부민병원(BUMIN HOSPITAL)0020687201106273228348'
    pw='bms819357*'
elif target=='hempel':
    target_title='헴펠_0000308533'
    pw=password
else:
    target_title='명유석(myoung yoosuk)0004014G000301779'
    pw=studiopass
# match_arrow_down(title)
# if driver.find_element_by_css_selector("[title^='" + title + "']") == None:
# if driver.find_element_by_xpath('//*[@title="주식회사밀앤아이(mill0225)0004029804094561"]') == None:
    # img_click('hometax05arrow')
    # pyautogui.move(10, None)
    # print('None')
    # pyautogui.press('down')
# else:
#     print('else')
    # driver.find_element_by_css_selector("[title^='" + title + "']").click()


# title = driver.find_element_by_xpath("//title") # 문서내의 어떤 태그든지 가능
#
# # head 태그 안에 있는 title 정보는 get_attribute('text') 메서드로 추출할 수 있습니다.
# print (title.get_attribute('text'))
# img_click('hometax04certicon')
# pyautogui.move(None, 100)
# pyautogui.scroll(100)


# 테이블 구조 파악하기
# rows=len(driver.find_elements_by_xpath("//tbody/tr[1]/td/span"))
# rows=len(driver.find_elements_by_xpath("//tbody/tr"))
rows=len(driver.find_elements(By.XPATH, "//td[1]"))
# cols=len(driver.find_elements_by_xpath("//tbody/tr[1]/td/span[1][title]"))
# print(rows)
# print(cols)
# i = str(1)
# current_title=driver.find_elements_by_xpath("//tbody/tr[1]/td[1]")
# current_title=driver.find_elements_by_xpath("//tbody/tr['" + i + "']/td[1]")
# current_title=driver.find_elements_by_xpath("//tbody/tr/td[1]")

current_title=driver.find_elements(By.xpath, "//td[1]")
# i=0
for item in current_title:
    # print(item.text)
    if item.text == target_title:
        driver.find_element(By.CSS_SELECTOR, "[title^='" + str(target_title) + "']").click()
        # print("o")
    # else:
    #     i = i + 1
    #     print(i)
else:
    if rows>10:
        for count in range(3):
            driver.find_element(By.CSS_SELECTOR, "div#MLjquiScrollAreaDownverticalScrollBardataTable").click()
    # print(item.text)
    elif rows > 8:
        for count in range(2):
            driver.find_element(By.CSS_SELECTOR, "div#MLjquiScrollAreaDownverticalScrollBardataTable").click()
    elif rows > 6:
        driver.find_element(By.CSS_SELECTOR, "div#MLjquiScrollAreaDownverticalScrollBardataTable").click()
    driver.find_element(By.CSS_SELECTOR, "[title^='" + target_title + "']").click()
##################################################3
total1=driver.find_elements(By.CSS_SELECTOR, "tbody>tr>td>span[title]")
# total1=driver.find_elements_by_css_selector("tbody>tr>id")
# total1=driver.find_element_by_xpath('//*tbody/tr/td/span')
# totalCount = driver.find_element_by_class_name('_fd86t').text
# total1=driver.find_elements_by_css_selector("tbody>tr>id").text
# print("총 게시물:", total1)
# total1=driver.find_element_by_xpath("//table[@class='tabledataTable']/tbody/tr/td[1]/span[title]")
# total1=driver.find_element_by_xpath("//td/span[title]")


# for i in total1:
#     print(i.text)
#     if i.text == target_title:
#         driver.find_element_by_xpath("[title^='" + target_title + "']").click()
#         # driver.find_element_by_css_selector("[title^='" + target_title + "']")
#         break
# else:
#     for k in range(10):
#         pyautogui.press('down')
#     # pyautogui.scroll(2000)
#     total2 = driver.find_elements_by_css_selector("span[title]")
#     for j in total2:
#         print(j.text)
#         if j.text == target_title:
#             # driver.find_element_by_xpath("[title^='" + j.text + "']").click()
#             driver.find_element_by_css_selector("[title^='" + j.text + "']").click()
#             break



# SCROLL_PAUSE_TIME = 0.5
#
# while True:
#
#     # Get scroll height
#     ### This is the difference. Moving this *inside* the loop
#     ### means that it checks if scrollTo is still scrolling
#     last_height = driver.execute_script("return document.body.scrollHeight")
#
#     # Scroll down to bottom
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#
#     # Wait to load page
#     sleep(SCROLL_PAUSE_TIME)
#
#     # Calculate new scroll height and compare with last scroll height
#     new_height = driver.execute_script("return document.body.scrollHeight")
#     if new_height == last_height:
#
#         # try again (can be removed)
#         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#
#         # Wait to load page
#         sleep(SCROLL_PAUSE_TIME)
#
#         # Calculate new scroll height and compare with last scroll height
#         new_height = driver.execute_script("return document.body.scrollHeight")
#
#         # check if the page height has remained the same
#         if new_height == last_height:
#             # if so, you are done
#             break
#         # if not, move on to the next loop
#         else:
#             last_height = new_height
#             continue
#######################################################
# driver.find_element_by_xpath('//*[@title="주식회사밀앤아이_0000413351"]').click()
# driver.find_element_by_xpath('//*[@title="헴펠_0000308533"]').click()
# driver.find_element_by_css_selector('tr#row0dataTable').click()

# for elem in driver.find_elements_by_xpath('.//span[@title = "주식회사밀앤아이_0000413351"]'):
#     print(elem.text)

# target_title='주식회사밀앤아이_0000413351'
# match_arrow_down(title)
# if driver.find_element_by_css_selector("[title^='" + title + "']") == None:
# if driver.find_element_by_xpath('//*[@title="주식회사밀앤아이(mill0225)0004029804094561"]') == None:
    # img_click('hometax05arrow')
    # pyautogui.move(10, None)
    # print('None')
    # pyautogui.press('down')
# else:
#     print('else')
    # driver.find_element_by_css_selector("[title^='" + title + "']").click()


# title = driver.find_element_by_xpath("//title") # 문서내의 어떤 태그든지 가능
#
# # head 태그 안에 있는 title 정보는 get_attribute('text') 메서드로 추출할 수 있습니다.
# print (title.get_attribute('text'))
# img_click('hometax04certicon')
# pyautogui.move(None, 100)
# # pyautogui.scroll(100)
# for i in range(2):
#     driver.find_element_by_css_selector("div#MLjquiScrollAreaDownverticalScrollBardataTable").click()
#
# total1=driver.find_elements_by_css_selector("span[title]")
# for i in total1:
#     # print(i.text)
#     if i.text == target_title:
#         driver.find_element_by_xpath("[title^='" + target_title + "']").click()
#         # driver.find_element_by_css_selector("[title^='" + target_title + "']")
#         break
# else:
#     for k in range(10):
#         pyautogui.press('down')
#     # pyautogui.scroll(2000)
#     total2 = driver.find_elements_by_css_selector("span[title]")
#     for j in total2:
#         print(j.text)
#         if j.text == target_title:
#             # driver.find_element_by_xpath("[title^='" + j.text + "']").click()
#             driver.find_element_by_css_selector("[title^='" + j.text + "']").click()
#             break
##############################################




# current_title = driver.find_element_by_css_selector("[title]").text
# print(current_title)
# driver.find_element_by_xpath("[title^='" + current_title + "']").click()
# driver.find_element_by_css_selector("[title^='" + current_title + "']").click()
# print("title :", current_title)
# for i in current_title:
#     print("title :", i)
# pyautogui.press('down')
# sleep(3)




# def match_mouse_scroll(cert):
#     while True:
#         img_icon = pyautogui.locateOnScreen(img_filename)
#         if img_icon == None:
#             pyautogui.scroll(-5)
#         else:
#             img_icon_center = pyautogui.center(img_icon)
#             # pyautogui.click(img_icon_center)
#             pyautogui.doubleClick(img_icon_center)
#             break

# num = 0
# for i in range(10):
#     pyautogui.press('down')
#     num = +1
# driver.find_element_by_css_selector("[title^='" + title + "']").click()




# driver.find_element_by_css_selector("[title^='"+title+"']").click()
sleep(3)
# 패스워드
# driver.find_element_by_css_selector("input#input_cert_pw").click()
driver.find_element(By.ID, 'input_cert_pw').send_keys(pw)

# 확인
driver.find_element(By.ID, 'btn_confirm_iframe').click()
sleep(5)
# 사업장
if target =='studio':
    driver.find_element(By.ID, 'btnApply').click()
    sleep(3)
    iframe = driver.find_element(By.CSS_SELECTOR, "iframe#UTXPPAAA24_iframe")
    driver.switch_to.frame(iframe)
    sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#grid1_cell_2_0 > input").click()
    sleep(3)
    driver.find_element(By.CSS_SELECTOR, "#trigger47").click()
    sleep(3)
    if driver.switch_to.alert != None:
        alert = driver.switch_to.alert
        alert.accept()
    sleep(5)
    if driver.switch_to.alert != None:
        alert = driver.switch_to.alert
        alert.accept()



# img_click('hometax02hdd')
# sleep(3)
# img_click('hometax03hdd_G')

# img_mouse_scroll('hometax04_mill')
# img_arror_click('hometax04_mill')
# num = 0
# for i in range(10):
#     pyautogui.press('down')
#     num = +1

# img_arror_down('hometax04_mill')
# img_mouse_scroll('hometax04_mill')
#img_arror_click('hometax05certpass')
# img_click('hometax05certpass')
# pyperclip.copy("패스워드")
# pyautogui.hotkey("ctrl", "v")
# pyautogui.typewrite(password)
# pyautogui.press('enter')

sleep(3)
driver.find_element(By.CSS_SELECTOR, "a#group1300.w2group").click()
sleep(3)

# 광고창 닫기
parent_window=driver.current_window_handle
# driver.execute_script("window.open('https://www.naver.com')")
all_windows=driver.window_handles
child_window=[window for window in all_windows if window != parent_window][0]
driver.switch_to.window(child_window)
driver.title
driver.close()
driver.switch_to.window(parent_window)
driver.title

iframe=driver.find_element(By.CSS_SELECTOR, "html>body>div#wrap>div#container_wrap>iframe#txppIframe")
driver.switch_to.frame(iframe)
sleep(3)
# driver.find_element_by_css_selector("ul#group1608.w2group.sub_layer01.sub_layer07>a#a_0104020000.w2textbox").click()
driver.find_element(By.CSS_SELECTOR, "a#a_0104020000.w2textbox").click()
sleep(3)
driver.find_element(By.CSS_SELECTOR, "a#a_0104020300.w2textbox").click()
sleep(3)

# 전자세금계산서
driver.find_element(By.CSS_SELECTOR, "input#radioEtxivClsfCd_input_0").click()
sleep(5)

# 매입
driver.find_element(By.CSS_SELECTOR, "input#radio3_input_1").click()
# sleep(3)

# 조회기간 반기별
driver.find_element(By.CSS_SELECTOR, "input#radio4_input_2").click()
# sleep(3)

# 조회년도
select=Select(driver.find_element(By.ID, 'selectboxYear'))
select.select_by_visible_text('2019')
# sleep(3)

# 기수 선택
select=Select(driver.find_element(By.ID, 'selectboxQrt2'))
select.select_by_visible_text('2기')
# sleep(3)

# 정렬 선택
select=Select(driver.find_element(By.ID, 'srtOpt'))
select.select_by_visible_text('오름차순')
# sleep(3)

# 조회하기 버튼
driver.find_element(By.CSS_SELECTOR, "input#trigger50").click()
sleep(5)

# 엑셀내려받기
driver.find_element(By.ID, "trigger55").click()
sleep(3)

# 파일종류선택확인
# iframe=driver.find_element_by_css_selector("html>body>div#wrap>div#container_wrap>iframe#UTEETBDA17_iframe")
iframe=driver.find_element(By.CSS_SELECTOR, "iframe#UTEETBDA17_iframe")
driver.switch_to.frame(iframe)

driver.find_element(By.CSS_SELECTOR, "input#btnProcess").click()
sleep(5)
###################################
# option_text=driver.find_element_by_css_selectors("select#crrnPageForExcelDwlld>option")
option_text=driver.find_elements(By.XPATH, "//option")
# for item in option_text:
#     print(item.text)
# print(option_text[0].text)
option_count=len(driver.find_elements(By.XPATH, "//option"))
# print(option_count)

# i=0
# for i in range(3):
#     option_text[i].click()
#     sleep(1)

# print(driver.window_handles)
path="c:/Users/miero/downloads/temp/"
newpath="c:/Users/miero/downloads/hometax/"
# path="c:/Users/miero/downloads/hometax/"
# newpath="c:/Users/miero/downloads/hometax/"

cName="hometax_"
i=0
for i in range(option_count):
    option_text[i].click()
    driver.find_element(By.CSS_SELECTOR, "input#trigger4").click()
    # getDownLoadedFileName(5)
    sleep(2)
    # alert 창의 '확인'을 클릭합니다.
    if driver.switch_to.alert != None:
        alert = driver.switch_to.alert
        alert.accept()
    if i ==1:
        img_click('hometax_multifile_ok')
    sleep(5)
    for filename in os.listdir(path):
        print(path + filename, '=>', newpath + str(cName) + str(i+1) + '.xls')
        os.rename(path + filename, newpath + str(cName) + str(i+1) + '.xls')
        i = i + 1
####################################################
# 취소
driver.find_element(By.CSS_SELECTOR, "#trigger10004").click()
sleep(5)
iframe=driver.find_element(By.CSS_SELECTOR, "html>body>div#wrap>div#container_wrap>iframe#txppIframe")
driver.switch_to.frame(iframe)
sleep(5)

# 전자계산서
driver.find_element(By.CSS_SELECTOR, "#radioEtxivClsfCd_input_1").click()

sleep(3)
# 매입
driver.find_element(By.CSS_SELECTOR, "#radio3_input_1").click()
sleep(3)
# 조회기간 반기
driver.find_element(By.CSS_SELECTOR, "#radio4_input_2").click()
# sleep(3)

# 조회년도
select=Select(driver.find_element(By.ID, 'selectboxYear'))
select.select_by_visible_text('2019')
# sleep(3)

# 기수 선택
select=Select(driver.find_element(By.ID, 'selectboxQrt2'))
select.select_by_visible_text('2기')
# sleep(3)

# 정렬 선택
select=Select(driver.find_element(By.ID, 'srtOpt'))
select.select_by_visible_text('오름차순')
# sleep(3)

# 조회하기 버튼
driver.find_element(By.CSS_SELECTOR, "input#trigger50").click()
sleep(5)

# 없으면 로그아웃
# if driver.switch_to.alert != None:
#     alert = driver.switch_to.alert
#     alert.accept()
print(driver.find_element(By.CSS_SELECTOR, "#txtTotal").text)
if driver.find_element(By.CSS_SELECTOR, "#txtTotal").text == '0':
    # iframe = driver.find_element_by_css_selector("#txppIframe")
    # driver.switch_to.frame(iframe)
    # driver.find_element_by_css_selector("a#group1332.w2group ").click()
    # sleep(3)
    # if driver.switch_to.alert != None:
    #     alert = driver.switch_to.alert
    #     alert.accept()
    driver.close()

# 엑셀내려받기
driver.find_element(By.ID, "trigger55").click()
sleep(3)

# 파일종류선택확인
# iframe=driver.find_element_by_css_selector("html>body>div#wrap>div#container_wrap>iframe#UTEETBDA17_iframe")
iframe=driver.find_element(By.CSS_SELECTOR, "iframe#UTEETBDA17_iframe")
driver.switch_to.frame(iframe)

driver.find_element(By.CSS_SELECTOR, "input#btnProcess").click()
sleep(5)
###################################
# 테이블 구조 파악하기
# rows=len(driver.find_elements_by_xpath("//tbody/tr[1]/td/span"))
# rows=len(driver.find_elements_by_xpath("//tbody/tr"))
# rows=len(driver.find_elements_by_xpath("//td[1]"))

###################################
# option_text=driver.find_element_by_css_selectors("select#crrnPageForExcelDwlld>option")
option_text=driver.find_elements(By.XPATH, "//option")
# for item in option_text:
#     print(item.text)
# print(option_text[0].text)
option_count=len(driver.find_elements(By.XPATH, "//option"))
# print(option_count)

# i=0
# for i in range(3):
#     option_text[i].click()
#     sleep(1)

# print(driver.window_handles)
path="c:/Users/bm115/downloads/temp/"
newpath="c:/Users/bm115/downloads/hometax/"
# path="c:/Users/miero/downloads/hometax/"
# newpath="c:/Users/miero/downloads/hometax/"
cName="hometaxfree_"
i=0
for i in range(option_count):
    option_text[i].click()
    driver.find_element(By.CSS_SELECTOR, "input#trigger4").click()
    # getDownLoadedFileName(5)
    sleep(2)
    # alert 창의 '확인'을 클릭합니다.
    if driver.switch_to.alert != None:
        alert = driver.switch_to.alert
        alert.accept()
    sleep(5)
    for filename in os.listdir(path):
        print(path + filename, '=>', newpath + str(cName) + str(i+1) + '.xls')
        os.rename(path + filename, newpath + str(cName) + str(i+1) + '.xls')
        # i = i + 1
# 취소
driver.find_element(By.CSS_SELECTOR, "#trigger10004").click()
sleep(5)
iframe=driver.find_element(By.CSS_SELECTOR, "#txppIframe")
driver.switch_to.frame(iframe)

driver.find_element(By.CSS_SELECTOR, "#group1332").click()
sleep(5)
if driver.switch_to.alert != None:
    alert = driver.switch_to.alert
    alert.accept()
    # for filename in os.listdir():
    #
    #     # 파일 확장자가 (properties)인 것만 처리
    #     # if filename.endswith("properties"):
    #         # 파일명에서 AA를 BB로 변경하고 파일명 수정
    #     new_filename = filename.replace("AA", "BB")
    #         os.rename(filename, new_filename)

    # latestDownloadedFileName = getDownLoadedFileName(10) #waiting 3 minutes to complete the download
    # # sleep(5)
    # print(latestDownloadedFileName)
    # # sleep(2)
    # driver.close()
    # sleep(2)
    # shutil.copy("c:/Users/miero/Downloads/"+latestDownloadedFileName,"c:/Users/miero/Downloads/hometax/hometax" + str(i) + ".xls")
    # driver.switch_to.window(driver.window_handles[0])
    # sleep(5)

    # print(driver.window_handles)




# option_text[0].click()
# driver.find_element_by_css_selector("input#trigger4").click()
# # getDownLoadedFileName(3)
# sleep(2)
# # alert 창의 '확인'을 클릭합니다.
# alert = driver.switch_to.alert
# alert.accept()
# # img_click('hometax_download_ok')
# # sleep(5)
# latestDownloadedFileName = getDownLoadedFileName(10) #waiting 3 minutes to complete the download
# # sleep(5)
# print(latestDownloadedFileName)
# # sleep(2)
driver.close()
# sleep(2)
# shutil.copy("c:/Users/miero/Downloads/"+latestDownloadedFileName,'c:/Users/miero/Downloads/hometax/hometax1.xls')
# i=str(1)
# pyautogui.typewrite('taxinvoice' + i + '.xls')
#
#
# pyautogui.press('tab')
# pyautogui.press('tab')
# pyautogui.press('tab')
# pyautogui.press('tab')
# pyautogui.press('tab')
#
# pyautogui.typewrite('C:\\Users\\miero\\hometax')
# pyautogui.press('enter')
# sleep(2)

# driver.execute_script("window.open()")










# sleep(2)
# for i in range(option_count):
#     option_text[i].click()
#     # print(option_text[i].text)
#     driver.find_element_by_css_selector("input#trigger4").click()
#     sleep(3)
#     getDownLoadedFileName(3)
#     # img_click('hometax_download_ok')










# NEXT_BUTTON_XPATH = '//input[@type="button" and @id="trigger38" and @class="w2trigger "]'
# button = driver.find_element_by_xpath(NEXT_BUTTON_XPATH)
# button.click()
# # 공인인증서 영역
# iframe = driver.find_element_by_id('dscert')
# driver.switch_to.frame(iframe)
#
# # 두번째 공인인증서 선택
# driver.find_element_by_xpath('//*[@title="주식회사밀앤아이(mill0225)0004029804094561"]').click()
#

# driver.find_element_by_id('input_cert_pw').send_keys(passwd)
#
# driver.find_element_by_id('btn_confirm_iframe').click()

# 상위 프레임으로 이동
# driver.switch_to_default_content()

# driver.close()