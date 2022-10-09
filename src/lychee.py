from selenium import webdriver
from selenium.webdriver.chrome.service import Service # load file browser 
from selenium.webdriver.chrome.options import Options # tùy biến e chọn profile thì e dùng cái này
from selenium.webdriver.common.by import By # để tìm element của trang web 
from selenium.webdriver.common.keys import Keys # 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
def getInfo():
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
       a.send_keys(i[-1])
       a.send_keys(Keys.ENTER)


getInfo()

# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[6]/div/strong  TEAM
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[3]/div  RESOURCE
# /html/body/app-root/div/div[1]/div/div[2]/app-user-management/section/div/div[2]/app-users-table/div/div/div/div[1]/div[2]/div[1]/div[7]/div ROLE
# class = no-data