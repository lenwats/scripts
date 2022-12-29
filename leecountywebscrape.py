import openpyxl

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException

import time

options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option('detach', True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

url = "https://www.leepa.org/Search/PropertySearch.aspx"
driver.get(url)
time.sleep(1)
street_input = driver.find_element(By.ID, "ctl00_BodyContentPlaceHolder_WebTab1_tmpl0_AddressTextBox")      
  
wrkbk = openpyxl.load_workbook("LeeCountyClaims.xlsx")
  
sheet = wrkbk.active

row = sheet.max_row
column = sheet.max_column


for i in range(1073, 2001): 
    driver.get(url)

    street_input = driver.find_element(By.ID, "ctl00_BodyContentPlaceHolder_WebTab1_tmpl0_AddressTextBox")
    
    time.sleep(1)

    cell = sheet.cell(row = i, column = 18) 
    street_input.send_keys(cell.value)
    street_input.send_keys(Keys.ENTER)

    results = WebDriverWait(driver, 15).until(ec.presence_of_element_located((By.XPATH, "/html/body/form/div[5]/div/div/div[1]/div/div/div/div[1]/div[1]/table/tbody/tr/td[1]/div/div[2]"))).text
    
    folio_url = "https://www.leepa.org/Display/DisplayParcel.aspx?FolioID="
    full_url = f"{folio_url}{results}"
    print(full_url)


    # ** used for navigating to the parcel details page and printing the whole url - probs unnecessary ** #
    #results = driver.find_element(By.XPATH, "/html/body/form/div[5]/div/div/div[1]/div/div/div/div[1]/div[1]/table/tbody/tr/td[4]/div/div[1]/a").click()
    #driver.switch_to.window(driver.window_handles[-1])
    #print(driver.current_url)


