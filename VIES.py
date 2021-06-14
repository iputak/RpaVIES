import openpyxl
import os
import json

from selenium import webdriver
from selenium.webdriver.support.ui import Select

def excel():
    path = 'excel/vatNumbers.xlsx' 
    wb = openpyxl.load_workbook(path) 
    wsMemberState = wb.get_sheet_by_name('Member State')
    wsRequestMemberState = wb.get_sheet_by_name('Request Member State')
    wsResult = wb.get_sheet_by_name('Results')
    brRedova = wsRequestMemberState.max_row
    url = wsMemberState['A'+'2'].value
    memberStateCodeImport = wsMemberState['B'+'2'].value
    vatNumberImport = wsMemberState['C'+'2'].value
    for i in range(2, brRedova+1): 
        requesterMemberStateCodeImport = wsRequestMemberState['A'+str(i)].value
        requesterNumberImport = wsRequestMemberState['B'+str(i)].value
        consultationNumbers = web_login(url, memberStateCodeImport, vatNumberImport, requesterMemberStateCodeImport, requesterNumberImport)
        wsResult.append ([requesterNumberImport, requesterMemberStateCodeImport, consultationNumbers])
        wb.save(path)

def web_login(url, memberStateCodeImport, vatNumberImport, requesterMemberStateCodeImport, requesterNumberImport):
    path_to_file = os.getcwd()+'\VIES' 
    chrome_options = webdriver.ChromeOptions()
    settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
        "download.default_directory": path_to_file,
        'savefile.default_directory': path_to_file
    }

    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    CHROMEDRIVER_PATH = './drivers/chromedriver'
    driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)
    driver.get(url)

    memberStateCode = Select(driver.find_element_by_name("memberStateCode"))
    vatNumber = driver.find_element_by_name("number")
    requesterMemberStateCode = Select(driver.find_element_by_name("requesterMemberStateCode"))
    requesterNumber = driver.find_element_by_name("requesterNumber")
    verifyButton = driver.find_element_by_name("check")

    memberStateCode.select_by_value(memberStateCodeImport)
    vatNumber.send_keys(vatNumberImport)
    requesterMemberStateCode.select_by_value(requesterMemberStateCodeImport)
    requesterNumber.send_keys(requesterNumberImport)
    verifyButton.click() 

    consultationNumbers = driver.find_element_by_css_selector('tbody tr:nth-child(8) td:nth-child(2)').text
    
    if consultationNumbers == "" : 
        consultationNumbers = "//"
    
    driver.execute_script('document.title="{}";'.format(str(requesterMemberStateCodeImport)+str(requesterNumberImport)));        
    driver.execute_script('window.print();')
    driver.close()
    driver.quit()
    return consultationNumbers
excel()