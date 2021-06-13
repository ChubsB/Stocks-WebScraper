from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

url = "http://www.scstrade.com/stockscreening/SS_CompanySnapShotHP.aspx?symbol=power"
driver = webdriver.Chrome(ChromeDriverManager().install())

driver.get(url)
content = driver.page_source 
soup = BeautifulSoup(content, "lxml")

#test = soup.findAll('tr', class_ = "jqgrow ui-row-ltr")
date = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[1]").text
openx = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[2]").text
high = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[3]").text
low  = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[4]").text
close = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[5]").text
volume = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[2]/td[6]").text
close_previous = driver.find_element_by_xpath("/html[1]/body[1]/form[1]/div[4]/div[2]/div[2]/div[2]/div[1]/div[1]/div[3]/div[1]/ul[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[3]/div[1]/table[1]/tbody[1]/tr[3]/td[5]").text
change = round(float(close) - float(close_previous), 2)
print(date, openx, high, low, close, volume,change)

