from typing import Dict, Any, Union

from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from bs4 import BeautifulSoup
import time
import pyautogui
from openpyxl import load_workbook

baseUnitOfMeasureDic = {'1/2 Pallet' : 'HPL',
                        '1/1 Pallet': 'DP',
                        '1/6 Pallet': 'CA',
                        'Pallet': 'PL',
                        'Case' : 'CA',
                        '1/4 Pallet' : 'QPL',
                        'Display': 'DP',
                        'Floorstand': 'DP',
                        'Piece': 'PC',
                        '1/8 Pallet': 'CA'}
SKUTypeDic = {
         'regular' : 'T1',
         'promotional' : 'T2',
         'industrial' : 'T3',
         'commercial': 'T4',
         'mto/unique pallets' : 'T7',
        'co-packing input': 'T8',
        'None': 'T1'}

asVendorDic= {
    'NL52 - KUEHNE&NAGEL NL': 'NL52',
    'NL52':'NL52',
    'PL50 - FM Olszowa' :   'PL50',
    'RO50 - K&N' :   'RO50',
    'DE51 - Druck + Pack' :   'DE51',
    'CH50 - Planzer' :   'CH50',
    'FR57 - FM Crépy' :   'FR57',
    'DK50 - Frode Laursen' :   'DK50',
    'GB50 - DHL' :   'GB50',
    'GR20 - Diakinisis' :   'GR20',
    'ID logisitics' :   'ES50',
    'PL20 - Sonoco' : 'PL20',
    'FRG6 - ALLOGA' :'FRG6',
    'FRG5 - RAFFIN' :'FRG5',
    'ES50-Logiters' :'ES50',
    'ES50 - POLO' : 'ES50',

}

dataDic= {
    'PL79': 'PL20',
    'PL78': 'PL20',
    'CZ99': 'PL50',
    'HU99': 'PL50',
    'PL99': 'PL50',
    'SK98': 'PL50',
    'RO99': 'RO50',
    'AT99': 'DE51',
    'DE99': 'DE51',
    'CH99': 'CH50',
    'BE99': 'NL52',
    'NL99': 'NL52',
    'FR99': 'FR57',
    'DK99': 'DK50',
    'FI99': 'DK50',
    'NO99': 'DK50',
    'SE99': 'DK50',
    'GB99': 'GB50',
    'GR99': 'GR20',
    'IT99': 'IT50',
    'PT99': 'ES50',
    'ES99': 'ES50'
}


languagesDic = {
    'BE99' : ['French', 'Dutch','German'],
    'PL78' : ['Polish'],
    'PL79' : ['Polish'],
    'PL99' : ['Polish'],
    'NL99' : ['Dutch'],
    'CZ99' : ['Czech'],
    'RO99' : ['Romanian'],
    'AT99' : ['German'],
    'DE99' : ['German'],
    'CH99' : ['German', 'Italian','French'],
    'FR99' : ['French'],
    'DK99' : ['Danish'],
    'FI99' : ['Finnish'],
    'NO99' : ['Norwegian'],
    'SE99' : ['Swedish'],
    'GR99' : ['Greek'],
    'IT99' : ['Italian'],
    'PT99' : ['Portuguese'],
    'ES99' : ['Spanish'],
    'GB99' : ['English_ZA']
}

InnovationDic={
'NO' : ['NO - Base Business'],
'H1' : ['H1 - Core Innovation'],
'H1R' : ['H1R - Core Relaunch'],
'H2' : ['H2 - Novel'],
'H3' : ['H3 - Breakthrough'],
'GEO' : ['GEO - Geographic Expansion']

}

sapUrl = 'http://appportal.win.colpal.com/irj/portal?NavigationTarget=navurl://551876c535632da41f46ac6ddd9ccb22'

sessionFile = open("sessionIdAndProduct.txt", "r")
sessionAndExecutorUrl = sessionFile.read().splitlines()
session_id = sessionAndExecutorUrl[0]
executor_url = sessionAndExecutorUrl[1]

def attach_to_session(sap_session_id, executor_url):
    original_execute = WebDriver.execute
    def new_command_execute(self, command, params=None):
        if command == "newSession":
            return {'success': 0, 'value': None, 'sessionId': sap_session_id}
        else:
            return original_execute(self, command, params)
    WebDriver.execute = new_command_execute
    driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
    driver.session_id = sap_session_id

    WebDriver.execute = original_execute
    return driver

driver = attach_to_session(session_id, executor_url)
driver.execute_script('''window.open("http://bings.com","_blank");''')


for window in driver.window_handles:
    driver.switch_to.window(window)
    if 'Bing' in driver.title:
        break
# print(driver.current_window_handle)
driver.get(sapUrl)
# print('attached')
# print(driver.current_window_handle)

# print(driver.window_handles)
workbook = load_workbook(filename="ProjectData.xlsx")
sheet1 = workbook.active
templateData = {
	"projectNumber": sheet1["A2"].value,
    "ExpShipDate": sheet1["B2"].value,
    "ProductMix":sheet1["C2"].value,
    "ProjectCreator": sheet1["D2"].value,
    "MatDesc":sheet1["E2"].value,
    "LocalMatDesc":sheet1["F2"].value,
    "SalesOrg":sheet1["G2"].value,
    "ReplacingSKU" :sheet1 ["H2"].value,
    "BaseSKU" :sheet1["I2"].value,
    "UOM" : sheet1["J2"].value,
    "EANStrategy" :sheet1["K2"].value,
    "Type": sheet1["L2"].value,
    "PPG": sheet1["M2"].value,
    "AsVendor" : sheet1["N2"].value
}

MDG = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tabIndex2")))
MDG.click()
time.sleep(5)

materialRequests = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "L2N6")))
materialRequests.click()
time.sleep(5)



processMaterial = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "6L3N0")))
processMaterial.click()
time.sleep(5)

SKUCreate = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "60L4N0")))
SKUCreate.click()

time.sleep(2)
driver.switch_to.frame('contentAreaFrame')

# DESCRIPTION
isolatedWorkAreaFrame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//iframe[@name='isolatedWorkArea']")))
driver.switch_to.frame(isolatedWorkAreaFrame)
descriptionInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table/tbody/tr[3]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[3]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr[1]/td[2]/div/table/tbody/tr[3]/td[3]/table/tbody/tr/td/input')))
descriptionInput.click()
descriptionInput.send_keys(("MIH" if "PL20" in templateData["AsVendor"] else "CPL") + ' ' + str(templateData["MatDesc"]).strip() + templateData["projectNumber"])
time.sleep(5)

#NOTES
notes = driver.find_elements_by_class_name('lsTbsv5-ItemTitle')[1]
notes.click()
time.sleep(3)
newNote = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table/tbody/tr[3]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[3]/td/table/tbody/tr[3]/td/div[2]/div/div/div/table/thead/tr/th/table/tbody/tr/td/span/div')))
newNote.click()
time.sleep(3)
driver.switch_to.default_content()
newNoteFrame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//iframe[@name="URLSPW-0"]')))
print(newNoteFrame)
driver.switch_to.frame(newNoteFrame)
newNoteInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td[2]/textarea')))
newNoteInput.send_keys(templateData['projectNumber'])
newNoteOk = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/div/div[1]/div/div[4]/div/table/tbody/tr/td[3]/table/tbody/tr/td[1]/div')
newNoteOk.click()
time.sleep(3)
driver.switch_to.default_content()
driver.switch_to.frame('contentAreaFrame')
driver.switch_to.frame(isolatedWorkAreaFrame)
time.sleep(3)

#SUBSIDIARY

subsidiaryInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[4]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/span')))
subsidiaryInput.click()

subsidiaryInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[4]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/input')))
subsidiaryInput.send_keys('EP')

assignButton = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[4]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')))
assignButton.click()
time.sleep(2)




#SOURCING TYPE
sourcingTypeInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@title="Sourcing Type" and @class="lsField__input"]')))
sourcingTypeInput.send_keys('Manufactured In-house' if "PL20" in templateData["AsVendor"] else 'Copacked Local')
time.sleep(2)


#BASE UNIT OF MEASURE
baseUnitInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]/div/table/tbody/tr[2]/td[3]/table/tbody/tr/td[1]/input')))
baseUnitInput.click()
baseUnitInput.click()
baseUnitInput.send_keys(baseUnitOfMeasureDic[templateData['UOM']])
time.sleep(1)


#SALES ORGS
salesOrgValues = str(templateData["SalesOrg"]).strip().split()

if "PL20" in templateData["AsVendor"]:
    salesOrgAdd = salesOrgValues.append('PL79')


salesOrg = driver.find_elements_by_class_name('lsTbsv5-ItemTitle')[1]
salesOrg.click()

for i, salesOrgValue in enumerate(salesOrgValues):
    time.sleep(3)
    addRow = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')))
    addRow.click()


    url = None
    if (i != 0):
        url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(i + 1) + ']/td[1]/div/span/span'
    else:
        url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/span'
    salesOrgInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, url)))
    salesOrgInput.click()

    time.sleep(2)

    if (i != 0):
        url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(i + 1) + ']/td[1]/div/span/input'
    else:
        url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/input'
    salesOrgInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, url)))
    salesOrgInput.click()

    salesOrgInput.send_keys(salesOrgValue)
    salesOrgInput.send_keys(Keys.RETURN)
    time.sleep(3)



    SKUType = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                              '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/div/table/tbody/tr[3]/td[6]/table/tbody/tr/td[1]/input')))
    SKUType.click()
    SKUTypeValue = str(templateData['Type']).lower()
    SKUType.send_keys(SKUTypeDic[str(templateData['Type']).lower()])
    time.sleep(2)

    Innovation = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                 '/ html / body / table / tbody / tr / td / div / div[1] / table / tbody / tr[2] / td / table / tbody / tr / td / div / div[1] / table / tbody / tr[5] / td / div / table / tbody / tr / td / table / tbody / tr[2] / td / div / div / table / tbody / tr[2] / td / table / tbody / tr[3] / td / div[2] / div / div / table / tbody / tr[2] / td / table / tbody / tr[4] / td[3] / table / tbody / tr / td[1] / input')))
    Innovation.click()
    Innovation.send_keys('NO - Base Business')
    time.sleep(2)

    # if InnovationValue != '':
    #
    #     Innovation = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
    #                                                                                  '/ html / body / table / tbody / tr / td / div / div[1] / table / tbody / tr[2] / td / table / tbody / tr / td / div / div[1] / table / tbody / tr[5] / td / div / table / tbody / tr / td / table / tbody / tr[2] / td / div / div / table / tbody / tr[2] / td / table / tbody / tr[3] / td / div[2] / div / div / table / tbody / tr[2] / td / table / tbody / tr[4] / td[3] / table / tbody / tr / td[1] / input')))
    #     Innovation.click()
    #     InnovationValue = str(templateData['Innovation']).lower()
    #     Innovation.send_keys(InnovationDic[str(templateData['Innovation']).lower()])
    #     time.sleep(2)



    materialGroup1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                     '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/div/table/tbody/tr[2]/td[3]/table/tbody/tr/td[1]/input')))
    materialGroup1.click()
    materialGroup1.send_keys('0OH')
    materialGroup1.send_keys(Keys.RETURN)
    time.sleep(2)

    materialGroup3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/div/table/tbody/tr[4]/td[3]/table/tbody/tr/td[1]/input')))
    materialGroup3.click()
    materialGroup3.send_keys('002' if SKUTypeValue == 'regular' else '003')
    materialGroup3.send_keys(Keys.RETURN)
    time.sleep(1)

    #MATERIAL GROUP 4
    PPGValue = str(templateData['PPG'])[0:3]
    if templateData['PPG'] == 'PPG will be filled in MDG':
         PPGValue= ''

    if PPGValue != '':

        materialGroup4 = WebDriverWait(driver, 10).until((EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[2]/div/div/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/div/table/tbody/tr[5]/td[3]/table/tbody/tr/td[2]/span'))))
        materialGroup4.click()
        time.sleep(2)
        driver.switch_to.default_content()
        iframeGroup4 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//iframe[@name='URLSPW-0']")))
        driver.switch_to.frame(iframeGroup4)
        addSearch = WebDriverWait(driver, 10).until((EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div[1]/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span[3]/a/span'))))
        addSearch.click()
        localValueField= WebDriverWait(driver, 10).until((EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div[1]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[5]/td[3]/table/tbody/tr/td/input'))))
        localValueField.click()
        localValueField.send_keys(PPGValue)
        localValueField.send_keys(Keys.RETURN)
        time.sleep(1)
        fieldEnter = WebDriverWait(driver, 10).until((EC.presence_of_element_located((By.XPATH,'/html/body/table/tbody/tr/td/div/div[1]/div/div[3]/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td[1]/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[3]/div'))))
        fieldEnter.click()
        time.sleep(3)


    driver.switch_to.default_content()
    driver.switch_to.frame('contentAreaFrame')
    driver.switch_to.frame(isolatedWorkAreaFrame)

#PLANT

if  templateData["AsVendor"] != "PL20 - Sonoco" :
    asVendor = driver.find_elements_by_class_name('lsTbsv5-ItemTitle')[2]
    asVendor.click()
    plantValue = str(asVendorDic[templateData['AsVendor']])
    print(plantValue)

    addRow = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')))
    addRow.click()

    asVendorInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/span ')))

    asVendorInput.click()
    asVendorInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/input')))
    asVendorInput.click()

    asVendorInput.send_keys(plantValue)
    asVendorInput.send_keys(Keys.RETURN)
    time.sleep(5)
    countryOfOrigin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                      '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[3]/table/tbody/tr/td[1]/input')))
    countryOfOrigin.click()
    countryOfOrigin.click()
    countryOfOriginValue = str(asVendorDic[templateData['AsVendor']])[:2]
    countryOfOrigin.send_keys(countryOfOriginValue)
    countryOfOrigin.send_keys(Keys.RETURN)

else:
    asVendorValues = str(asVendorDic[templateData['AsVendor']]).strip().split()
    uniquePlants = set()


    for i, salesOrgValue in enumerate(salesOrgValues):
        uniquePlants.add(dataDic[salesOrgValue])


    asVendor= driver.find_elements_by_class_name('lsTbsv5-ItemTitle')[2]
    asVendor.click()


    for i, asVendorValue in enumerate(uniquePlants):
        time.sleep(3)
        addRow = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                             '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')))
        addRow.click()

        url = None
        if (i != 0):
            url ='/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(i + 1) + ']/td[1]/div/span/span'
        else:
            url ='/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/span'
        asVendorInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, url)))
        asVendorInput.click()

        time.sleep(2)
        if (i != 0):
            url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(i + 1) + ']/td[1]/div/span/input'
        else:
            url = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[1]/td/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[1]/div/span/input'


        asVendorInput = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, url)))
        asVendorInput.click()

        asVendorInput.send_keys(asVendorValue)
        asVendorInput.send_keys(Keys.RETURN)
        time.sleep(5)


        countryOfOrigin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                                          '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[5]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[3]/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[3]/table/tbody/tr/td[1]/input')))
        countryOfOrigin.click()
        countryOfOrigin.click()
        countryOfOriginValue = str(asVendorDic[templateData['AsVendor']])[:2]
        countryOfOrigin.clear()
        countryOfOrigin.send_keys(countryOfOriginValue)
        countryOfOrigin.send_keys(Keys.RETURN)




#Additional Data

time.sleep(3)
additionalData = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/div/div/span[1]')))
additionalData.click()

additionalDataDesc = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[2]/div/span/span')))
additionalDataDesc.click()
additionalDataDesc = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr/td[2]/div/span/input')))
additionalDataDesc.click()
additionalDataDesc.send_keys(str(templateData["MatDesc"]).strip())
additionalDataDesc.send_keys(Keys.RETURN)

time.sleep(1)



addRow = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
                                                                             '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')))
print(salesOrgValues)

uniqueLanguages = set()

for i, salesOrgValue in enumerate(salesOrgValues):
    for language in languagesDic[salesOrgValue]:
        uniqueLanguages.add(language)

for j, language in enumerate(uniqueLanguages):
    time.sleep(3)
    addRow = driver.find_element_by_xpath(
        '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/thead/tr/th/table/tbody/tr/td/span[1]/div')
    addRow.click()

    localMatDescXpath = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(
        (j + 2)) + ']/td[2]/div/span/span'
    localMatDesc = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, localMatDescXpath)))
    localMatDesc.click()

    localMatDescXpath = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(
        (j + 2)) + ']/td[2]/div/span/input'
    localMatDesc = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, localMatDescXpath)))
    localMatDesc.click()
    localMatDesc.send_keys(str(templateData['LocalMatDesc']).strip())


    languageXpath = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(
        (j + 2)) + ']/td[1]/div/table/tbody/tr/td[1]/span'
    languageInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, languageXpath)))
    languageInput.click()
    languageXpath = '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[6]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td/div[1]/div/div/table/tbody/tr/td/div/table/tbody[1]/tr[2]/td[2]/div/div[2]/table/tbody/tr[' + str(
        (j + 2)) + ']/td[1]/div/table/tbody/tr/td[1]/input'
    languageInput = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, languageXpath)))
    languageInput.click()
    languageInput.send_keys(language)
    languageInput.send_keys(Keys.RETURN)


#BUSINESS PARAMETERS
time.sleep(3)
businessParameters = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/div/div/span[2]')))
businessParameters.click()
projectNotes = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/div/table/tbody/tr[2]/td[3]/textarea')))
projectNotes.click()
time.sleep(2)
projectNotes = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/div/table/tbody/tr[2]/td[3]/textarea')))
projectNotesValue = str(templateData["MatDesc"].strip() +'\n' + "Expected Ship Date: " + templateData["ExpShipDate"].strip() + '\n' + "Product Mix: " + templateData["ProductMix"] + '\n'+ "ProjectCreator: " + templateData["ProjectCreator"].strip() + '\n' + "Base SKU:" + templateData["BaseSKU"].strip())
projectNotesValue = projectNotesValue.replace('\t', '        ')
print(projectNotesValue)
projectNotes.send_keys(projectNotesValue)


regularProductCheckBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[1]/td/table/tbody/tr[9]/td[2]/div/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/span')))
specialPackCheckBox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[1]/td/table/tbody/tr[9]/td[2]/div/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[2]/span')))
print(SKUTypeValue)
if (SKUTypeValue == 'regular'):
    regularProductCheckBox.click()
else:
    specialPackCheckBox.click()

yesPromotional = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[5]/td/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[1]/span')))
noPromotional = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr[7]/td/div/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/table/tbody/tr[5]/td/table/tbody/tr[2]/td[3]/table/tbody/tr[1]/td[2]/span')))
if (SKUTypeValue =='promotional'):
    yesPromotional.click()
else:
    noPromotional.click()

exit(0)


