from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By

from selenium.webdriver.common.proxy import Proxy, ProxyType
import selenium.common.exceptions

import openpyxl
import time
from selenium.common.exceptions import TimeoutException

chrome_options=webdriver.ChromeOptions()


driver=webdriver.Chrome( executable_path =r'C:\Users\Ansab Waseem\Desktop\haseebproj\chromedriver.exe' )
wait=WebDriverWait(driver,30)
url='https://www.modaalabutik.com/en-yeniler?baslangic=0&bitis=900&siralama=&ps=901s'
#url='https://www.modaalabutik.com/en-yeniler'

driver.get(url)


html=driver.page_source
soup=BeautifulSoup(html,"html.parser")
size=[]
title=[]
sku=[]
ima=[]
si=[]
count=0
workbook=openpyxl.load_workbook(r'file location') 
sheet=workbook.get_sheet_by_name('Sheet1')
wait=WebDriverWait(driver,10)
for dat in soup.findAll("div",{"class":"col-md-4 col-sm-6 col-xs-6 col-6 prodItem"}):
	size=[]
	ima=[]
	count+=1
	
	driver.find_element_by_xpath('/html/body/div[6]/div/div/div[2]/div[2]/div[1]/div['+str(count)+']').click()
	
	wait.until(EC.presence_of_all_elements_located((By.ID, 'demo4carousel')))
	for siz in driver.find_elements_by_xpath('//*[@id="veri-formu"]/div/div[2]/div[7]/label'):
		si=[siz.text]
		size.append(si)
	sz="D"+str(count)
	sheet[sz].value=str(size)
	print(size)
	images = driver.find_elements_by_xpath('//*[@id="demo4carousel"]/li/a/img')
	for image in images:
		im=[image.get_attribute('src')]
		ima.append(im)

	sz="C"+str(count)
	sheet[sz].value=str(ima)	
	print(ima)
	
	sk=[driver.find_element_by_xpath('//*[@id="tab_1"]/table/tbody/tr[1]/td[2]').text]
	sz="B"+str(count)
	sheet[sz].value=str(sk)
	print(sk)

	tit=[driver.find_element_by_xpath('//*[@id="veri-formu"]/div/div[2]/div[1]').text]
	sz="A"+str(count)
	sheet[sz].value=str(tit)
	print(tit)
	for tr in driver.find_elements_by_xpath('//*[@id="tab_1"]/table'):
		tds = tr.find_elements_by_tag_name('td')
		des=[td.text for td in tds]
	sz="E"+str(count)
	sheet[sz].value=str(des)
	print(des)
		
	driver.find_element_by_xpath('/html/body/div[5]/a').click()
	time.sleep(4)
	
	
workbook.save(r'file location')

