from selenium import webdriver 
import time 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas 
import xlsxwriter
import datetime 
import os

df = pandas.read_excel('SCL_Linkovi.xlsx', index_col=0)
vals_list = df['Linkovi'].tolist()

option = Options()

option.add_argument("--disable-infobars")
option.add_argument("start-maximized")
option.add_argument("--disable-extensions")
option.add_argument("--disable-notification")

def checkAddToCart(urls):
    outWorkbook = xlsxwriter.Workbook("BugReport.xlsx")
    outSheet = outWorkbook.add_worksheet()
    outSheet.write("A1", "URL")
    outSheet.write("B1", "Bug Report")
    outSheet.write("C1", "Screenshot")
    brojac = 0
    for i in urls: 
        driver = webdriver.Chrome(chrome_options=option, executable_path='chromedriver.exe')
        driver.get(i)
        
        time.sleep(6)
        addToBag = driver.find_elements(By.CLASS_NAME, "add-to-cart__button")
        isExists = len(addToBag)
        if isExists != 0:
            print(f"Button 'Add To Bag' exists on {i}.")
        else: 
            print(f"Button 'Add To Bag' does not exist on {i}")
            brojac += 1
            outSheet.write(brojac, 0, i)
            outSheet.write(brojac, 1, "Add to bag does not exist on this page.")
            outSheet.write(brojac, 2, "Seek reference for bug in C:\\Users\\Korisnik005\\Documents\\Python files\\PBI")
            outWorkbook.close()
            path = 'C:\\Users\\Korisnik005\\Documents\\Python Files'
            DateString = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
            os.chdir(path)
            NewFolder = 'PBI_' + DateString
            os.makedirs(NewFolder)
            driver.save_screenshot(NewFolder+'/bugfoto.png')
            

checkAddToCart(vals_list)