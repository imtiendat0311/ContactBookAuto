from selenium import webdriver
from selenium.webdriver.chrome.service import Service # load file browser 
from selenium.webdriver.chrome.options import Options # tùy biến e chọn profile thì e dùng cái này
from selenium.webdriver.common.by import By # để tìm element của trang web 
from selenium.webdriver.common.keys import Keys # 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd

def getSign():
    option = Options()
    option.add_argument('user-data-dir=/Users/boo/Library/Application Support/Google/Chrome')
    option.add_argument('--profile-directory=Profile 1')
    service = Service(executable_path='/Users/boo/Desktop/ContactBookAuto/chromem1')
    url='https://app.contractbook.com/documents/shared/60f03ae1-a700-4f01-a338-d02bb8c0def1?sortBy=%21createdAt&state=signed'
    browser = webdriver.Chrome(service=service,options=option)
    browser.get(url)
    WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH,'//*[local-name()="svg"][@class="icon--euAXU iconSuccess--+P3Qj"]')))
    c = browser.find_element(By.XPATH,'//*[@id="documents-page"]/div[2]/div/div[1]/div[2]/div[2]/button[4]/span/div/span/span').text[1:-1]
    b = browser.find_elements(By.XPATH,'//*[local-name()="svg"][@class="icon--euAXU iconSuccess--+P3Qj"]')
    i=0
    while i<2:
        try:
            b = browser.find_elements(By.XPATH,'//*[local-name()="svg"][@class="icon--euAXU iconSuccess--+P3Qj"]')
            if(len(b)/2==int(c)):
                break
            xpathAddr='//*[@id="documents-list"]/div['+str(26+(25*i))+']/button'
            xpathDoc='//*[@id="documents-list"]/div['+str(1+(25*i))+']'
            a=WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH,xpathDoc)))
            d=WebDriverWait(browser,30).until(EC.element_to_be_clickable((By.XPATH,xpathAddr)))
            browser.execute_script("arguments[0].scrollIntoView(true);", d)
            browser.execute_script("arguments[0].click();",d)
            i=i+1
        except TimeoutException:
            break
    b = browser.find_elements(By.XPATH,'//*[local-name()="svg"][@class="icon--euAXU iconSuccess--+P3Qj"]')
    print("Actual number of available signed contract "+c)
    print("Number of signed contract can find "+str(len(b)/2))
    f={'Sign Status':[],'Sign Date':[],'Email':[]}
    for i in range(len(b)):
        if i%2==0:
            f['Sign Status'].append(b[-1-i].get_attribute("data-tip")[0:6])
            f['Sign Date'].append(b[-1-i].get_attribute("data-tip")[7:])
            f['Email'].append(b[-1-i].get_attribute("data-for")[7:])
    df = pd.DataFrame(f, columns = ['Sign Status', 'Sign Date','Email'])
    df.to_excel(r'/Users/boo/Desktop/testing.xlsx',index = False, header=True)

getSign()

