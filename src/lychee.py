from selenium import webdriver
from selenium.webdriver.chrome.service import Service # load file browser 
from selenium.webdriver.chrome.options import Options # tùy biến e chọn profile thì e dùng cái này
from selenium.webdriver.common.by import By # để tìm element của trang web 
from selenium.webdriver.common.keys import Keys # 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
def getInfo():
    listIndex = [2,3,4,6,7,8,9]
    listElement=[[],[],[],[],[],[],[]]
    option = Options()
    option.add_argument('user-data-dir=/Users/boo/Library/Application Support/Google/Chrome')
    option.add_argument('--profile-directory=Profile 1')
    service = Service(executable_path='/Users/boo/Desktop/ContactBookAuto/chromem1')
    service = Service(executable_path='/Users/boo/Desktop/ContactBookAuto/chromem1')
    url='https://lychee.esoft.com/'
    excel=pd.read_excel('/Users/boo/Desktop/testing.xlsx')
    browser = webdriver.Chrome(service=service,options=option)
    browser.get(url)
    WebDriverWait(browser,30).until(EC.element_to_be_clickable((By.XPATH,"/html/body/app-root/div/div[1]/div/div[1]/app-header/ul/li[2]/a"))).click()
    for i in excel.values:
       a= WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH,"/html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[1]/app-users-filter/div/div[5]/input")))
       a.send_keys(Keys.COMMAND+'a') # bôi đen 
       a.send_keys(Keys.DELETE)
       a.send_keys(i[2])
       a.send_keys(Keys.ENTER)
       time.sleep(1)
       if len(browser.find_elements(By.XPATH,"//*[@class='no-data']"))==1:
            for j in range(7):
                listElement[j].append("No data")
       else:
            for g in range(7):
                r = listIndex[g]
                if r != 6:
                    t= WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH,"/html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div["+str(r)+"]/div"))).text
                    print(t)
                    listElement[g].append(t)
                else:
                    t= WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH,"/html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[6]/div/strong"))).text
                    print(t)
                    listElement[g].append(t)
    indexColumn = ['USERNAME','RESOURCE','NAME','TEAM','ROLE','LOCATION','STATUS']
    for i in range(7):
        excel[indexColumn[i]]=listElement[i]
    excel.to_excel(r'/Users/boo/Desktop/testing.xlsx',index = False, header=True)




getInfo()

# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1] # NO DATA

# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[2]/div USERNAME
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[3]/div RESOURCE
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[4]/div NAME
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[6]/div/strong TEAM
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[7]/div ROLE
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[8]/div LOCATION
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[9]/div STATUS

# WebDriverWait(browser,10).until(EC.any_of(EC.presence_of_element_located((By.XPATH,"//*[@class='no-data']")),EC.element_to_be_clickable((By.XPATH,"//*[@class='table-td content-m']"))))
# lambda browser: browser.find_element(By.XPATH,"//*[@class='no-data']") or len(browser.find_elements(By.XPATH,"//*[@class='table-row']"
# time.sleep(2)