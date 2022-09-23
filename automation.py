from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import date
import time
import pandas as pd

# AUTHOR : Dat Nguyen with contribute by Khanh Linh

def autofill(d):
    b = "D:\ContactBookAuto\chromedriver.exe" # dir for chrome driver 
    option = Options()
    option.add_argument('user-data-dir=C:\\Users\\khanh\\AppData\\Local\\Google\\Chrome\\User Data') # chrome profile dir
    option.add_argument('--profile-directory=Default') # choose default profile
    service = Service(executable_path=b) # start chrome service
    browser = webdriver.Chrome(service=service,options=option)
    # URL = d[len(d)-1] # link url nho de o cuoi  || THis may not need 
    URL="https://app.contractbook.com/drafter/Esoft/freelancer-contract-7139e0f0/sign-in"
    browser.get(URL)
    time.sleep(5)
    #email field
    browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/span/div/div/input").send_keys("scintern.vn@esoft.com")
    time.sleep(2)
    #button to sign in
    browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/button").click()
    time.sleep(2)
    #get start button
    browser.find_element(By.XPATH,"/html/body/div[6]/div/div[1]/div/div/div/button").click()  
    time.sleep(2)
    d.insert(1,"Việt Nam") #country
    d.insert(3,"Vietnamese") #nationality 
    d.insert(4,d[2].split(", ")[-1]) #region chia address bằng dấu , rồi lấy phần cuối của address 
    listID = ['fullName','country','address','nationality','region','dateOfBirth',
    'passportOrNationalIdNumber','issuedDateAndUssuedPlaceOfPassportid',
    'accountNumber','bankName','bankAddress','swiftCode','email']
    # browser.find_element(By.ID,"lm-accept-necessary").click()
    # loop đi từ 0 -> 11 bỏ qua email tại vì email cần dùng lệnh riêng 
    for i in range(len(listID)-1):
        a=browser.find_element(By.ID,listID[i]) # tìm ô cần điền 
        browser.execute_script("arguments[0].scrollIntoView(true);", a)# scroll tới ô cần điền 
        a.send_keys(Keys.CONTROL+'a') # bôi đen 
        a.send_keys(Keys.BACKSPACE) # xoá 
        time.sleep(1)
        a.send_keys(d[i]) # điền dữ liệu từ list 
    time.sleep(1)
    #email
    a=browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div[1]/div/form/div[13]/div[2]/div/div/input")
    browser.execute_script("arguments[0].scrollIntoView(true);", a)
    a.send_keys(Keys.CONTROL+'a')
    a.send_keys(Keys.BACKSPACE)
    a.send_keys(d[-1])
    time.sleep(1)
    a.send_keys(Keys.ENTER)
    time.sleep(3)
    # see doc button 
    browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div/div/button").send_keys(Keys.ENTER)
    time.sleep(3)
    # open contractbook button 
    browser.find_element(By.XPATH,'/html/body/div[1]/div/main/div/div/nav/div[1]/div/button').click()
    time.sleep(5)
    # browser.find_element(By.ID,'microsoft-button').click()
    browser.quit()

excel= pd.read_excel("D:\ContactBookAuto\Contract-Linh.xlsx",sheet_name='Sheet2',header=None)# header = None tức là k bỏ dòng đầu ß
#name address dob id number id issued id issued place,  bank ac , bank name , bank address , swiftcode ,email
c=[1,34,3,7,8,10,27,28,31,30,14] 
for i in excel.values:
    b=[] # tạo list chứa dữ liệu cần dùng cho form 
    for j in c:
        if j ==10: # nếu j=10 => nơi cấp id 
            b[4]+=" at "+i[j] # tại b[4] đang có ngày cấp => thêm at và nơi cấp vào string 
        else:    
            if isinstance(i[j],date): # nếu đây là dữ liệu ngày tháng 
                i[j]=i[j].strftime("%d/%m/%Y") # đổi ra đúng format ngày tháng năm 
            b.append(i[j]) # add dữ liệu vào list b
    print(b) # hiển thị dữ liệu list b ra màn hình ( Debug purpose )
    autofill(b) # gọi function để auto fill form 
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