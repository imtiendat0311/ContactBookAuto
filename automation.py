from selenium import webdriver
from selenium.webdriver.chrome.service import Service # load file browser 
from selenium.webdriver.chrome.options import Options # tùy biến e chọn profile thì e dùng cái này
from selenium.webdriver.common.by import By # để tìm element của trang web 
from selenium.webdriver.common.keys import Keys # 
from datetime import date # đổi giờ thời gian dateTime()
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time # ngủ =))
import pandas as pd # import dữ liệu từ excel

# AUTHOR : Dat Nguyen with contribute by Khanh Linh

def autofill(d):
    b = "D:\ContactBookAuto\chromedriver.exe" # dir for chrome driver 
    option = Options()
    option.add_argument('user-data-dir=C:\\Users\\khanh\\AppData\\Local\\Google\\Chrome\\User Data') # chrome profile dir
    option.add_argument('--profile-directory=Default') # choose default profil
    service = Service(executable_path=b) # start chrome service
    browser = webdriver.Chrome(service=service,options=option)
    start = time.perf_counter()
    # URL = d[len(d)-1] # link url nho de o cuoi  || This may not need d =[1,2,3] len(d)=3 d[2]=3 d[0]=1 
    URL="https://app.contractbook.com/drafter/Esoft/freelancer-contract-7139e0f0/sign-in" # biến chứa string
    browser.get(URL)

    # time.sleep(5)
    #email field
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/span/div/div/input"))).send_keys("scintern.vn@esoft.com")
    # browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/span/div/div/input").send_keys("scintern.vn@esoft.com")
    # time.sleep(2)
    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/button"))).click()
    #button to sign in
    # browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div[1]/form/div/button").click()
    # time.sleep(6)
    #get start button
    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[6]/div/div[1]/div/div/div/button"))).click()
    # browser.find_element(By.XPATH,"/html/body/div[6]/div/div[1]/div/div/div/button").click()  
    # time.sleep(3)
    d.insert(1,"Việt Nam") #country
    d.insert(3,"Vietnamese") #nationality 
    d.insert(4,"Asia") # region

    # 2 hang manh , hoan kiem , ha noi 
    # d.insert(4,d[2].split(", ")[-1]) #region chia address bằng dấu , rồi lấy phần cuối của address 
    listID = ['fullName','country','address','nationality','region','dateOfBirth',
    'passportOrNationalIdNumber','issuedDateAndUssuedPlaceOfPassportid',
    'accountNumber','bankName','bankAddress','swiftCode','email']
    # browser.find_element(By.ID,"lm-accept-necessary").click()
    # loop đi từ 0 -> 11 bỏ qua email (cần dùng lệnh riêng)  
    for i in range(len(listID)-1):
        a=WebDriverWait(browser,20).until(EC.presence_of_element_located((By.ID,listID[i]))) # tìm ô cần điền 
        browser.execute_script("arguments[0].scrollIntoView(true);", a)# scroll tới ô cần điền 
        a.send_keys(Keys.CONTROL+'a') # bôi đen 
        a.send_keys(Keys.BACKSPACE) # xoá 
        a.send_keys(d[i])
         # điền dữ liệu từ list 
    #email
    a=browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div[1]/div/form/div[13]/div[2]/div/div/input")
    browser.execute_script("arguments[0].scrollIntoView(true);", a)
    a.send_keys(Keys.CONTROL+'a')
    a.send_keys(Keys.BACKSPACE)
    a.send_keys(d[-1])
    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div[1]/div/form/div[14]/button"))).click()
    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div/div/button"))).click()
    # see doc button 
    WebDriverWait(browser,20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div/main/div/div/nav/div[1]/div/button'))).click()
    current_url = browser.current_url
    WebDriverWait(browser, 20).until(EC.url_changes(current_url))

    # browser.find_element(By.XPATH,"/html/body/div[1]/div/main/div/div[2]/div/div/div/button").send_keys(Keys.ENTER)
    # time.sleep(3)
    # open contractbook button 
    # browser.find_element(By.XPATH,'/html/body/div[1]/div/main/div/div/nav/div[1]/div/button').click()
    # time.sleep(7) 
    # browser.find_element(By.ID,'microsoft-button').click()
    browser.quit()
    stop = time.perf_counter()
    print(f"\n\nTake "+str(stop-start)+" s to finish fill form\n\n")



#excel= pd.read_excel(r"C:\Users\khanh\Desktop\ContactBookAuto-main\lam hop dong.xlsx",sheet_name='29.9',header = None,dtype='object')# header = None tức là k bỏ dòng đầu ß
#name, address, dob, idnumber, idissueddate, id issued place,  bank ac , bank name , bank address , swiftcode ,email
#"C:\Users\khanh\Desktop\lam hop dong 30.9.xlsx"
excel= pd.read_excel(r"C:\Users\khanh\Desktop\01.Photo_Hồ sơ freelance (06.10).xlsx",sheet_name='Contractbook Data Input',dtype='object')
c=[0,1,2,3,4,5,6,7,8,9,10] 
for i in excel.values:
    # value = excel.values[i]
    b=[] # tạo list chứa dữ liệu cần dùng cho form 
    for j in c:
        if j == 5: # nếu j=10 => nơi cấp id 
             b[4]+=" at "+i[j] # tại b[4] đang có ngày cấp => thêm at và nơi cấp vào string
        else:
            if isinstance(i[j],str):
                i[j]=i[j].strip()
            if isinstance(i[j],date): # nếu đây là dữ liệu ngày tháng 
                i[j]=i[j].strftime("%d/%m/%Y") # đổi ra đúng format ngày tháng năm 
            b.append(i[j]) # add dữ liệu vào list b
    print(b)
    print("\n") # hiển thị dữ liệu list b ra màn hình ( Debug purpose )
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