from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import random
import os
import pandas as pd

def autofill(d):
    a = input("What platform are u using\n input 1 for windows 2 for macos: ")
    #your credentials here
   
    cwd = os.getcwd()
    b = "/chromem1" if int(a)==2 else "/chromedriver.exe"
    #need chromedriver for your current chrome version
    browser = webdriver.Chrome(cwd+b)
    
    URL = d[len(d)-1] # link url nho de o cuoi
    browser.get(URL)
    time.sleep(5)
    cookButton = browser.find_element(By.ID,'lm-accept-necessary')
    cookButton.click()
    # startButton =  browser.find_element(By.XPATH,'//button[normalize-space()="Start"]')
    # startButton.click()
    time.sleep(5)
    listID = ['fullName','country','address','nationality','region','dateOfBirth',
    'passportOrNationalIdNumber','issuedDateAndUssuedPlaceOfPassportid',
    'accountNumber','bankName','bankAddress','swiftCode','email']
    # inputFName = browser.find_element(By.ID,'fullName')
    # inputCountry = browser.find_element(By.ID,'country')
    # inputAddress = browser.find_element(By.ID,'address')
    # inputNationality = browser.find_element(By.ID,'nationality')
    # inputRegion = browser.find_element(By.ID,'region')
    # inputDob = browser.find_element(By.ID,'dateOfBirth')
    # inputID = browser.find_element(By.ID,'passportOrNationalIdNumber')
    # inputIssuedDate = browser.find_element(By.ID,'issuedDateAndUssuedPlaceOfPassportid')
    # inputACNumb= browser.find_element(By.ID,'accountNumber')
    # inputBank=browser.find_element(By.ID,'bankName')
    # inputBankAddr = browser.find_element(By.ID,'bankAddress')
    # inputSwift = browser.find_element(By.ID,'swiftCode')
    # inputEmail = browser.find_element(By.ID,'email')
    listField =[]
    for i in range(3):
        a=browser.find_element(By.ID,listID[i])
        a.send_keys(d[i])
    # browser.execute_script("browser.find_element(By.ID,'email').scrollIntoView();",browser.find_element(By.ID,'email'))
    # browser.find_element(By.ID,'email').send_keys("dat@gmail.com")
    # for i in listField:
    #     i.send_keys("a")
    #     time.sleep(2)
    # time.sleep(100)

    # #fill credentials and log in
    # userIdField = browser.find_element_by_id("fullName")
    # passwordField = browser.find_element_by_id("country")   
    # loginButton = browser.find_elements_by_tag_name("button")[1]
    # # userIdField.send_keys(userId)
    # # passwordField.send_keys(password)
    # loginButton.click()
    

#     #wait for two factor auth
#     time.sleep(5)
#     notNowButton = browser.find_element_by_tag_name("a")
#     notNowButton.click()
#     #should be able to get into post now
#     for i in postURL:
#         browser.get(i)
#         time.sleep(10)
#         # comment x many times
#         for j in cmt:
#             time.sleep(timedelta)
#             commentBox = browser.find_element_by_id("composerInput")
#             try:
#                 commentBox.send_keys(j)
#             except:
#                 #try 10 times max with 1 second delay if comment box is not available. 
#                 for i in range(0,10):
#                     try:
#                         time.sleep(1)
#                         commentBox.send_keys(j)
#                         break
#                     except:
#                         pass
#             #leave this delay fixed. Usually takes 2-3 seconds for the comment to get posted
#             time.sleep(10)
#             try:
#                 sendButton = browser.find_elements_by_tag_name("button")[0]
#                 sendButton.click()
#             except:
#                 #try 10 times max with 1 second delay if comment box is not available. 
#                 for i in range (0,10):
#                     try:
#                         time.sleep(1)
#                         sendButton = browser.find_elements_by_tag_name("button")[0]
#                         sendButton.click()
#                         break
#                     except:
#                         pass
            
        
# # list of post need to comment
# # list of post need to comment
# postUrl=["https://m.facebook.com/groups/1344873668901428/permalink/5170024369719653/",
#          "https://m.facebook.com/groups/330966831016688/permalink/1157555105024519/",
#         #  "https://m.facebook.com/groups/vieclamquanbinhtan/permalink/5146164988765265/", 
#          "https://m.facebook.com/groups/danbinhchanhnew/permalink/5067636606646783/",
#         # "https://m.facebook.com/groups/377156329334346/permalink/1611130729270227/",
#         "https://m.facebook.com/groups/pgpbkhuvucmienbac/permalink/425006882503944/",
#         "https://m.facebook.com/groups/101725677141334/permalink/982419969071896/",
#         # "https://m.facebook.com/groups/406678599486879/permalink/2337853006369419/",
#         # "https://m.facebook.com/groups/timvieclamtaihcmsaigon2/permalink/3147531062169326/",
#         # "https://m.facebook.com/groups/275568093015527/permalink/1154702545102073/",
#         "https://m.facebook.com/groups/280503298758830/permalink/2311721458970327/"
#         ]

# cmt = ["up"]
# call method ( how many repeate comment, how long of the delay comment, list of post )
# autofill()

excel= pd.read_excel("D:\Esoft\Contract - Linh.xlsx",sheet_name='Sheet1')
# list chua vi tri o A=0 B=1 .... ok a
c=[0,3,5,10,11,14]
for i in excel.values:
    b=[]
    for j in c:
        b.append(i[j])
    autofill(b)
print("finish")

# name 0 
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