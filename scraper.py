from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import openpyxl

#This is our Stocks list; if you want to add new stocks add it's name in this list
stocks_list = ['AKBL', 'ANL', 'ASC', 'ASL', 'ASTL', 'ATRL', 'AVN', 'BAFL', 'BIPL', 'BOP', 'BYCO', 'DCL', 'DOL', 'EFERT', 'EPCL', 'FABL', 'FCCL', 'FCSC', 'FFBL', 'FFC', 'FFL', 'FNEL', 'GATM', 'GGL', 'GGGL', 'HASCOL', 'HUMNL', 'ICIBL', 'ICL', 'ISL', 'JSBL', 'JSCL',
               'KAPCO', 'KEL', 'KOSM', 'LOADS', 'LOTCHEM', 'MDTL', 'MLCF', 'NBP', 'NRSL', 'PACE', 'PAEL', 'PIAA', 'PIBTL', 'PIOC', 'POWER', 'PPL', 'PRL', 'PSX', 'PTC', 'SILK', 'SNGP', 'SPL', 'SSGC', 'STCL', 'STPL', 'TELE', 'TPL', 'TREET', 'TRG', 'UNITY', 'WAVES', 'WTL']
stocks_list_test = ['KEL', 'ASC']
days = 19
days += 1

#This is to make sure the browser itself does not show 
options = Options()
options.headless = True
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)


for i in stocks_list_test:
    url = "http://www.scstrade.com/stockscreening/SS_CompanySnapShotHP.aspx?symbol=" + i
    driver.get(url)
    content = driver.page_source
    soup = BeautifulSoup(content, "lxml")

    for y in range(2, days+1):
        date = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[1]").text
        openx = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[2]").text
        high = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[3]").text
        low = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[4]").text
        close = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[5]").text
        volume = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y) + "]/td[6]").text
        close_previous = driver.find_element_by_xpath(
            "/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[" + str(y+1) + "]/td[5]").text
        change = round(float(close) - float(close_previous), 2)
        print(i+": ", date, openx, high, low, close, volume, change)
    driver.implicitly_wait(1)

driver.quit()
