from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import date
import time
import random
import os
import pandas as pd

def autofill(d):
    b = "/Users/boo/Desktop/ContactBookAuto/chromem1"
    browser = webdriver.Chrome(b)
    # URL = d[len(d)-1] # link url nho de o cuoi
    URL="https://app.contractbook.com/drafter/Esoft/freelancer-contract-7139e0f0?answersSetId=7bc020af-9782-4a6a-ac14-fe74f52a2445&step=0&source=Esoft"
    browser.get(URL)
    time.sleep(5)
    d.insert(1,"Việt Nam") #country
    d.insert(3,"Việt Nam") #region
    d.insert(4,d[2].split(", ")[-1])
    listID = ['fullName','country','address','nationality','region','dateOfBirth',
    'passportOrNationalIdNumber','issuedDateAndUssuedPlaceOfPassportid',
    'accountNumber','bankName','bankAddress','swiftCode','email']
    for i in range(len(listID)-1):
        a=browser.find_element(By.ID,listID[i])
        browser.execute_script("arguments[0].scrollIntoView(true);", a)
        a.send_keys(Keys.COMMAND+'a')
        a.send_keys(Keys.DELETE)
        time.sleep(2)
        a.send_keys(d[i])
    time.sleep(1)
    #email
    a=browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div[1]/div/form/div[13]/div[2]/div/div/input")
    browser.execute_script("arguments[0].scrollIntoView(true);", a)
    a.send_keys(Keys.COMMAND+'a')
    a.send_keys(Keys.DELETE)
    a.send_keys(d[-1])
    time.sleep(1)
    a.send_keys(Keys.ENTER)
    time.sleep(5)
    browser.quit()

excel= pd.read_excel("/Users/boo/Desktop/ContactBookAuto/Contract-Linh.xlsx",sheet_name='Sheet2',header=None)
#name address dob id number id issued bank ac , bank name , bank address , swiftcode ,email
c=[1,34,3,7,8,27,28,31,30,14]
for i in excel.values:
    b=[]
    for j in c:
        if isinstance(i[j],date):
            i[j]=i[j].strftime("%d/%m/%Y")
        b.append(i[j])
    print(b)
    autofill(b)
print("finish")

# name 1 
# dob 3 
# passport 7
# date issue 8
# place issue 10
# address 11
# email 14
# contract date ?? 20
# bank number 27
# bank name 28
# bank address 29
# swift code 30