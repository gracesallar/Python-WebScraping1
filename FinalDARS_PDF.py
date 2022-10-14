##This program downloads DARS from the Faculty and Advising Center.
##Written by Grace Sallar - Russ College of Engineering and Technology


import pandas as pd
import xlrd
from os import path
from pandas import ExcelWriter
from pandas import ExcelFile
from datetime import datetime
from selenium import webdriver
from getpass import getpass
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException 

file_path = path.relpath("Students.xlsx")
data = pd.read_excel(file_path)

PIDs = list(data.iloc[:,0])
CatYr = list(data.iloc[:,1])


choice = str(input('Would you like to update DARS? Type Yes or No: '))
##choice = 'Yes'

driver = webdriver.Chrome()
driver.get('https://cas.sso.ohio.edu/login?service=https%3A%2F%2Fwebapps.ohio.edu%2Foasis%2F')


wait = WebDriverWait(driver, 300);
wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="qm0"]/a[1]')))
total = -1

for x in PIDs:
    print("Looking up DARS for "+x)
    total = total + 1


    def check_request_audit_by_id(xID):
        try:
            driver.find_element_by_id(xID)
        except NoSuchElementException:
            return False
        return True

    def check_refresh_audit_by_date(xpath):
        try:
            date = driver.find_element_by_xpath(xpath).get_attribute('value')
        except NoSuchElementException:
            return False
        return date      
    
    driver.find_element_by_xpath('//*[@id="qm0"]/a[2]').click()
    driver.find_element_by_xpath('//*[@id="content"]/div/ul/li[3]/a').click()
    PID = driver.find_element_by_id('identityQuery')
    PID.click()
    PID.send_keys(x)
    driver.find_element_by_xpath('//*[@id="lookupForm"]/button[1]/span').click()

    request_audit = check_request_audit_by_id('button_'+CatYr[total]+'_SEMESTER')
    refresh_check = check_refresh_audit_by_date('//*[@id="semester_audit_'+CatYr[total]+'"]/form[1]/input[6]')

    def refresh_audit(xpath):
        if refresh_check[5:16] != datetime.today().strftime("%b %d %Y"):
            Year = driver.find_element_by_xpath('//*[@id="carYear_'+CatYr[total]+'"]').get_attribute('value')
            return True
        else:
            return False


    def IPDARS():
        try:
            ipdars = driver.find_element_by_class_name('ipDagger')
        except NoSuchElementException:
            return False
        return True    
        

    # This is for graduates

    if request_audit is True:   # Only request audit is available
        driver.find_element_by_id(str('button_'+CatYr[total]+'_SEMESTER')).click()
        wait.until(EC.presence_of_element_located((By.XPATH,(str('//*[@id="semester_audit_'+CatYr[total] + '"]/form[1]/input[6]')))))
        driver.find_element_by_xpath((str('//*[@id="semester_audit_'+CatYr[total] + '"]/form[1]/input[6]'))).click()
        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="qm0"]/li[5]/a')))
        driver.find_element_by_xpath('//*[@id="qm0"]/li[5]/a').click()



    elif choice == 'Yes' or choice == 'YES' or choice == 'Y' or choice == 'y' or choice == 'yes': # This refreshes audit
        Year = driver.find_element_by_xpath('//*[@id="catalogYear_'+CatYr[total]+'"]').get_attribute('value')

        if IPDARS() is True:
            try:
                driver.find_element_by_xpath('//*[@id="semester_audit_'+CatYr[total]+'"]/button').click()
                driver.find_element_by_xpath('//*[@id="RequestAuditButtonipRequest'+CatYr[total]+'SEMESTER"]/span').click()
                wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="semester_audit_'+CatYr[total]+'"]/button')))
            except NoSuchElementException:
                driver.find_element_by_xpath('//*[@id="darsRefreshForm'+CatYr[total]+'SEMESTER'+Year+'"]/input[5]').click()
                refresh_check = check_refresh_audit_by_date('//*[@id="semester_audit_'+CatYr[total]+'"]/form[1]/input[6]')
                wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="darsRefreshForm'+CatYr[total]+'SEMESTER'+Year+'"]/input[5]')))

        else:
            driver.find_element_by_xpath('//*[@id="darsRefreshForm'+CatYr[total]+'SEMESTER'+Year+'"]/input[5]').click()
            refresh_check = check_refresh_audit_by_date('//*[@id="semester_audit_'+CatYr[total]+'"]/form[1]/input[6]')
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="darsRefreshForm'+CatYr[total]+'SEMESTER'+Year+'"]/input[5]')))

    
        newbtn = driver.find_element_by_xpath((str('//*[@id="semester_audit_'+ CatYr[total] + '"]/form[1]/input[6]')))
        newbtn.click()               
        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="qm0"]/li[5]/a')))
        driver.find_element_by_xpath('//*[@id="qm0"]/li[5]/a').click()

    elif  choice == 'No' or choice == 'NO' or choice == 'N' or choice == 'n':                       

        driver.find_element_by_xpath((str('//*[@id="semester_audit_'+ CatYr[total] + '"]/form[1]/input[6]'))).click()        
        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="qm0"]/li[5]/a')))
        driver.find_element_by_xpath('//*[@id="qm0"]/li[5]/a').click()
                                         
    else:       # No need to refresh the audit                       

        driver.find_element_by_xpath((str('//*[@id="semester_audit_'+ CatYr[total] + '"]/form[1]/input[6]'))).click()        
        wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="qm0"]/li[5]/a')))
        driver.find_element_by_xpath('//*[@id="qm0"]/li[5]/a').click()

