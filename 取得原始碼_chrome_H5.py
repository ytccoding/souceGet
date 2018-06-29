# -*- coding: utf-8 -*-
from selenium import webdriver
from time import sleep
from openpyxl import workbook ,load_workbook ,Workbook
import os ,time ,ytFuntion
from bs4 import BeautifulSoup

def sheet_date():
    sheet["F" + str(i)].value = time.strftime("%y_%m_%d") #檢查日期
    sheet["G" + str(i)].value = time.strftime("%H_%M_%S") #檢查時間

testdayFile = time.strftime("%y_%m_%d")
testdayTime  = time.strftime("%y_%m_%d_%H_%M_%S")

if not os.path.exists(testdayFile + "_H5"):    #先確認資料夾是否存在
    os.makedirs(testdayFile + "_H5")
    
chrome_path = "D:\selenium_driver_chrome\chromedriver.exe" #webdriver放置資料夾
mobileEmulation = {'deviceName': 'iPhone 6/7/8'}
options = webdriver.ChromeOptions()
options.add_experimental_option('mobileEmulation', mobileEmulation)

webDriver = webdriver.Chrome(chrome_path ,chrome_options=options)
test_web = ytFuntion.test_web(webDriver)

wb = load_workbook("前台_H5.xlsx")
sheet = wb["web"] # 獲取一張表

testWebUrl = input("請輸入測試站點的url(Ex.http://m.fatkuma.com):")
sheet["A1"].value = testWebUrl
webDriver.get(str(testWebUrl) + str(sheet["D2"].value).strip())
print()

for i in range(2 ,len(sheet["B"])+1):
    if i == 10:
        test_web.elementSendKeys("input[tag=帐号]" ,6 ,text = "ytau1")
        test_web.elementSendKeys("input[tag=密码]" ,6 ,text = "qwe123")
        webDriver.execute_script("document.getElementsByTagName('a')[7].click()") #立即登入的點擊法
        sleep(5)
    testUrl = str(testWebUrl) + str(sheet["D" + str(i)].value).strip()
    sheet["D" + str(i)].value = testUrl
    webDriver.get(testUrl)
    sleep(3)
    text = open(os.getcwd() + "\\" + testdayFile + "_H5" + "\\" + str(i) + "_" + \
                str(sheet["B" + str(i)].value).strip() + str(time.strftime("%y_%m_%d_%H_%M_%S")) + \
                ".txt" ,"w+" , encoding = "utf8")
    if webDriver.current_url == str(testUrl):
        sheet["E" + str(i)].value = "有"
        text.write(BeautifulSoup(webDriver.page_source, 'html.parser').prettify())
        sheet_date()
    else:
        sheet["E" + str(i)].value = "無此url"
        sheet_date()
        
wb.save(os.getcwd() + "\\" + testdayFile + "_H5" + "\\" + "原始碼_chrome_H5_" + str(time.strftime("%y_%m_%d_%H_%M_%S") + ".xlsx"))
webDriver.quit()
