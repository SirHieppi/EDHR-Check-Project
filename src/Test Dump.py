####################################################
# By: Daniel Brown (dbrown5)
#
# Comment: This file keeps all code I have tested in order to fix the eDHR checker.
####################################################

#Added here 23 JUL 2021
#This code was meant to open NC's in Fiori. However, I discoved that Fiori has a URL that can do this more easily
#so now this code is garbage.
    #Inputs: NC Number
    NCNumpath = '//a[contains(text(), "{}")'.format(str(300000011))
    
    #Open Fiori Manage NC tab
    FioriURL = 'https://sapfioriprd.illumina.com/sap/bc/ui5_ui5/ui2/ushell/shells/abap/FioriLaunchpad.html?sap-client=100&sap-language=EN#ZP2D_NC_MANAGEQN_SM-manage'
    chromedriver = r"\\ushw-file\Users\transfer\edhr_checker\eDHR\eDHR\src\chromedriver_win32\chromedriver.exe" 
    chromeOptions = webdriver.ChromeOptions()
    driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)
    driver.get(FioriURL)

    WaitForFiori()
    
    #Select "Open NC selection", see if NC exists here
    OpenNC = driver.find_elements_by_xpath('//div[@class="sapMSLITitleOnly"]')[1]
    OpenNC.click()
    
    WaitForFiori()
    
    try:
        NCrow = driver.find_elements_by_xpath(NCNumpath)
    except NoSuchElementException:
        print('NC 300000011 not here')




#Finished testing this code 19 JUL 2021
#The website Justin's code uses no longer lists the correct deviations
#Most of this code is copied over from "edhr.parse_mes_transactions_ruo"
#I should try finding the deviations listed in the MES list
#Also, I need to find if a deviation was excluded for any reason

#"edhr.parse_deviations" is the code that parses deviations; it returns a dictionary; see below for example
#{'OP750': ['1046251'], 'OP100': ['1041393'], 'OP300': ['1046251'], 'OP400': ['1046251'], 'OP500': ['1046251'], 'OP550': ['1046251'], 'OP600': ['1046251']}

try:
    PATH_NAME = dirname(dirname(abspath(__file__)))
    SN='M07500'
    xl = pd.ExcelFile(PATH_NAME + "/eDHR_Files/{}.xlsx".format(SN))
    xlFile = xl
    csv = open(PATH_NAME + '\eDHR_Files\{}.csv'.format(SN), mode='r', encoding='utf8')
    
    row_offsets = edhr.parse_csv(csv)
    
    #Run the parse_deviations function
    devns = edhr.parse_deviations(xl, row_offsets)
    print(devns)
    
    #New codethat looks through each line of the eDHR for deviations
    
    df = xlFile.parse(skiprows=row_offsets[4], nrows=row_offsets[5] - row_offsets[4] - 2, usecols="A,C,G")
    # Cols |   A   |   C   |    G    |
    keys = ['Step', 'Task', 'Result']
    
    for row in df.iterrows():
        temp = dict(zip(keys, row[1].values))
    
        if temp['Task'] == 'List Deviations' and not str(temp['Result']) == 'nan':
            op = str(temp['Step']).split('-')[0]
            deviation = str(temp['Result'])
            
            if op in devns:
                devns[op].append(deviation)
            else:
                devns[op] = [deviation]

    #Delete any duplicate values within a key
    for op in devns:
        devns[op] = list(set(devns[op]))
    
    print(devns)

    input('Press Enter to close window.')

    
except Exception:
    traceback.print_exc()
    sleep (60)    

#Testing dhr._parse_mes_transactions
#Finished testing this code 30 JUN 2021. Found out that I made the "material_num" variable an integer instead
#of a string. Changing it to a string fixed it.

try:

    PATH_NAME = dirname(dirname(abspath(__file__)))

    SN = 'M07600'
    ilm = "ilm"
    xl = pd.ExcelFile(PATH_NAME + "/eDHR_Files/{}.xlsx".format(SN))
    csv = open(PATH_NAME + '\eDHR_Files\{}.csv'.format(SN), mode='r', encoding='utf8')

    PLATFORM_TRANSACTIONS_MAP = {
        '15033616': edhr.parse_mes_transactions_ruo,
        '15048976': edhr.parse_mes_transactions_ruo,
        '20013740': edhr.parse_mes_transactions_nova,
        '20014737': edhr.parse_mes_transactions_dx
    }

    row_offsets = edhr.parse_csv(csv)
    material_num = '15033616'
    ilm_numbers = support.get_ilm_numbers(ilm)

    bad_transactions, fio_data_points, comments = PLATFORM_TRANSACTIONS_MAP[material_num](xl, row_offsets, dx_ilm_numbers if material_num == '20014737' else ilm_numbers)
    print('bad_transactions: {} \nfio_data_points: {} \ncomments: {}'.format(bad_transactions, fio_data_points, comments))
    sleep(60)


#This section of code is from edhr.parse_csv
#I tested it on 29 JUN 2021. It worked and returned a list of row offsets

import os
import csv
import sys
import time
from time import sleep
from itertools import chain
from os.path import dirname, abspath, isfile
import pandas as pd
from xlsxwriter.workbook import Workbook

import re
import math
from os.path import dirname, abspath
import pandas as pd
import numpy as np

try:
    PATH_NAME = dirname(dirname(abspath(__file__)))
    SN='M07600'
    csv = open(PATH_NAME + '\eDHR_Files\{}.csv'.format(SN), mode='r', encoding='utf8')

    #from edhr.parse_csv

    row_offsets = []

    with csv as file:
        count = 0
        for line in file.readlines():
            count += 1
            line = line.strip(',').strip()
            if not line:
                row_offsets.append(count)
    print (row_offsets)#return row_offsets\

    sleep(60)
    
except Exception as e:
    print ("\nError!")
    print (e)
    sleep (60)




# This section of code is from "downloadDHR" function found in "support.py"
# I finished testing this section on 24 JUN 2021. I updated the "support.py" code with edited sections
# namely, the "Configure Chrome profile to download to this script's directory"

try:
    
    #ask user for instrument serial number
    SN = input("Instrument serial number: ")
    
    #Define "pathName" variable
    PATH_NAME = dirname(dirname(abspath(__file__))) #The result for this is \\ushw-file\Users\transfer\edhr_checker\eDHR\eDHR
    pathName = PATH_NAME
    eDHR_FilesPath = "{}\eDHR_Files".format(pathName)

    if isfile(eDHR_FilesPath + "/{}.csv".format(SN)):
    	print('\n\tWARNING: eDHR file already exists! Skipping download. . .')
    	#return

    #CHANGE NEEDED
    # Configure Chrome profile to download to this script's directory
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : eDHR_FilesPath}
    chrome_options.add_experimental_option("prefs",prefs)
    chromedriver = "C:/bin/chromedriver.exe" #Will need to change this to something on the shared drive
    chrome_options.add_argument('--log-level=3')
    driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chrome_options)
    chrome_options.binary_location = r"C:/Program Files/Google/Chrome/Application/chrome.exe"

    # Open Chrome browser to eDHR website
    url = "http://ussd-illuminareporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fProduction%2feDHR+-+Instrument+-+Detail"
    #url = "http://ussd-illuminadevreporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fTest+Env%2feDHR+-+Instrument+-+Detail&SelectedSubTabId=GenericPropertiesTab"
    driver.get(url)

    # Enter Serial Number into form
    containerNameInput = driver.find_element_by_xpath('//*[@id="ctl32_ctl04_ctl03_txtValue"]')
    containerNameInput.send_keys(SN)
    containerNameInput.send_keys(Keys.ENTER)
    time.sleep(3)

    # Export and save
    saveIcon = driver.find_element_by_xpath('//*[@id="ctl32_ctl05_ctl04_ctl00_ButtonImgDown"]')
    saveIcon.click()
    time.sleep(1)
    exportCSV = driver.find_element_by_xpath('//*[@id="ctl32_ctl05_ctl04_ctl00_Menu"]/div[2]/a')
    exportCSV.click()
    time.sleep(2)


    # Error handling for when:
    #   (1) If eDHR files already exist in eDHR_Files folder, remove the newly downloaded file and proceed
    #   (2) If xlsx file is open and write permissions are not granted, terminate the script and prompt message
    try: 
        # Rename eDHR files to match SN
        csvToExcel(SN)
        os.rename("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath), "{}/{}.csv".format(eDHR_FilesPath, SN))
    except FileExistsError:
        print('\n\tWARNING: eDHR file already exists! Removing download. . .')
        os.remove("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath))
    except PermissionError:
        # Occurs when eDHR files already exist and the xlsx file is open. When it is open, it prevents
        # the writer from having write access. In this case, we no longer have to write and can ignore the
        # exception.
        print('\n\tWARNING: eDHR file already exists! Removing download. . .')
        os.remove("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath))

    driver.close()
    
    #CHANGE NEEDED
    #Get rid of the print an sleep
    print("eDHR Download Successful")
    sleep (3)


except Exception as e:
    print ("Error!")
    print (e)
    sleep (60)


'''
try:

    PATH_NAME = dirname(dirname(abspath(__file__))) #The result for this is \\ushw-file\Users\transfer\edhr_checker\eDHR\eDHR

    SN = 'm07600;
    #Export and download eDHR file for given SN. Save that file in two formats (CSV, XLSX)
    # If file exists, skip download
    eDHR_FilesPath = "{}/eDHR_Files".format(pathName)

    if isfile(eDHR_FilesPath + "/{}.csv".format(SN)):
    	print('\n\tWARNING: eDHR file already exists! Skipping download. . .')
    	return

    # Configure Chrome profile to download to this script's directory
    chrome_options = Options()
    prefs = {'download.default_directory' : eDHR_FilesPath}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--log-level=3')
    driver = webdriver.Chrome(options=chrome_options)

    # Open Chrome browser to eDHR website
    url = "http://ussd-illuminareporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fProduction%2feDHR+-+Instrument+-+Detail"
    #url = "http://ussd-illuminadevreporting.illumina.com/Reports/Pages/Report.aspx?ItemPath=%2fCamstar+Reports%2fTest+Env%2feDHR+-+Instrument+-+Detail&SelectedSubTabId=GenericPropertiesTab"
    driver.get(url)

    # Enter Serial Number into form
    containerNameInput = driver.find_element_by_xpath('//*[@id="ctl32_ctl04_ctl03_txtValue"]')
    containerNameInput.send_keys(SN)
    containerNameInput.send_keys(Keys.ENTER)
    time.sleep(3)

    # Export and save
    saveIcon = driver.find_element_by_xpath('//*[@id="ctl32_ctl05_ctl04_ctl00_ButtonImgDown"]')
    saveIcon.click()
    time.sleep(1)
    exportCSV = driver.find_element_by_xpath('//*[@id="ctl32_ctl05_ctl04_ctl00_Menu"]/div[2]/a')
    exportCSV.click()
    time.sleep(2)


    # Error handling for when:
    #   (1) If eDHR files already exist in eDHR_Files folder, remove the newly downloaded file and proceed
    #   (2) If xlsx file is open and write permissions are not granted, terminate the script and prompt message
    try: 
        # Rename eDHR files to match SN
        csvToExcel(SN)
        os.rename("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath), "{}/{}.csv".format(eDHR_FilesPath, SN))
    except FileExistsError:
        print('\n\tWARNING: eDHR file already exists! Removing download. . .')
        os.remove("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath))
    except PermissionError:
        # Occurs when eDHR files already exist and the xlsx file is open. When it is open, it prevents
        # the writer from having write access. In this case, we no longer have to write and can ignore the
        # exception.
        print('\n\tWARNING: eDHR file already exists! Removing download. . .')
        os.remove("{}/eDHR - Instrument - Detail.csv".format(eDHR_FilesPath))

    driver.close()
    print("Download Successful")
    sleep (10)

except Exception as e:
    print (e)
    sleep (60)
'''
    
'''
#This text is from "get_ncr_numbers" function in the "support.py" file. I am modifying it to work with Fiori.
#Comment: I am able to get it to open the "Manage NC" tile. However, I will be better off finding the NC number in the eDHR rather than here
try:

    chrome_options = Options()
    chrome_options.add_argument('--log-level=3')
    driver = webdriver.Chrome('//ushw-file/Users/transfer/edhr_checker/eDHR/eDHR/src/chromedriver_win32/chromedriver.exe')
    url = "https://sapfioriprd.illumina.com/sap/bc/ui5_ui5/ui2/ushell/shells/abap/FioriLaunchpad.html?sap-client=100&sap-language=EN#Shell-home"
    #"ETQ$CMD=CMD_OPEN_HOMEPAGE&ETQ$SCREEN_WIDTH=1280"
    driver.get(url)
    sleep (8)
    
    #Open NC webpage
    ncButton = driver.find_element_by_id('__tile144')
    ncButton.click()
    
    # Login
    inputUsername = driver.find_element_by_name("_FIELD--USER_NAME")
    inputPassword = driver.find_element_by_name("_FIELD--PASSWORD")
    inputUsername.send_keys(username)
    inputPassword.send_keys(password)
    inputPassword.send_keys(Keys.ENTER)
    
except Exception as e:
    print (e)
    sleep (60)
'''

'''
#This was to test if ChromeDriver works. (Change "from time import sleep" to "import time") Results: Program ran as intended.
try:
    driver = webdriver.Chrome('//ushw-file/Users/transfer/edhr_checker/eDHR/eDHR/src/chromedriver_win32/chromedriver.exe')  # Optional argument, if not specified will search path.

    driver.get('http://www.google.com/');

    time.sleep(5) # Let the user actually see something!

    search_box = driver.find_element_by_name('q')

    search_box.send_keys('ChromeDriver')

    search_box.submit()

    time.sleep(5) # Let the user actually see something!

    driver.quit()

except Exception as e:
    print (e)
    sleep (60)
'''

'''
#the following are the variables the function asks for in the original program
SN = 'M07600'
username = 'dbrown5'
password = 'abc'

try:
    # Open browser and navigate to ETQ Reliance
    chrome_options = Options()
    chrome_options.add_argument('--log-level=3')
    driver = webdriver.Chrome(options=chrome_options)
    url = "https://ussd-prd-etqr02.illumina.com/relianceprod/reliance?"
    "ETQ$CMD=CMD_OPEN_HOMEPAGE&ETQ$SCREEN_WIDTH=1280"
    driver.get(url)
    
    # Login
    inputUsername = driver.find_element_by_name('_FIELD--USER_NAME')
    inputPassword = driver.find_element_by_name('_FIELD--PASSWORD')
    inputUsername.send_keys(username)
    inputPassword.send_keys(password)
    inputPassword.send_keys(Keys.ENTER)
    
    # Open the Nonconformance tab
    ncIcon = driver.find_element_by_xpath('//*[@id="NCMR"]/div')
    ncIcon.click()

# If invalid login credentials, user will enter it again, or program will timeout
except NoSuchElementException:
    try:
        print("\n\tWARNING: Incorrect username or passsword. Please type in the correct username or password and proceed.")
        ncIcon = WebDriverWait(driver, 30).until(lambda driver: driver.find_element_by_xpath('//*[@id="NCMR"]/div'))
        ncIcon.click()
    except TimeoutException:
        print("\n\tWARNING: Program timed out due to inactivity")
        sys.exit(0)
except WebDriverException:
    sys.exit(0)

#trying to catch error message (not part of original code)
except exception as e:
    print (e)
    sleep (60)

# Search by Insturment Serial Number
openCloseToggle = driver.find_element_by_xpath('//*[@id="5._Open_and_Closed_8"]')
displayIndicator = driver.find_element_by_xpath('//*[@id="div_5._Open_and_Closed_8"]')
if displayIndicator.get_attribute("style") == "display: none;":
    openCloseToggle.click()
serialNumberButton = driver.find_element_by_xpath('//*//*[@id="BY_INSTRUMENT_SERIAL_NUMBER_2"]')
serialNumberButton.click()
searchTextBoxSN = driver.find_element_by_xpath('//*[@id="ColumnSearchPanel"]/td[2]/span/input')
searchTextBoxSN.send_keys(SN)
searchIcon = driver.find_element_by_xpath('//*[@id="SubmitColumnSearchIcon"]/img')
searchIcon.click()
time.sleep(2)

# Get NCR numbers from table
ncrNumbers = {}
numbers = driver.find_elements_by_xpath('//*[@id="mCSB_1_container"]/table/tbody/tr/td[3]')
materialNumbers = driver.find_elements_by_xpath('//*[@id="mCSB_1_container"]/table/tbody/tr/td[9]')
phases = driver.find_elements_by_xpath('//*[@id="mCSB_1_container"]/table/tbody/tr/td[5]')
for num, mn, phase in zip(numbers, materialNumbers, phases):
    ncrNumbers[num.get_attribute('textContent')] = (mn.get_attribute('textContent'), phase.get_attribute('textContent'))
driver.close()
return ncrNumbers
'''

'''
#These tests are to learn more about os.path commands and other commands
#print (Options())
#print (dirname(dirname(abspath(__file__))))
#print (type(dirname(dirname(abspath(__file__)))))
#print (dirname(__file__))
#print (abspath(__file__))
'''

'''
#This test ensures the program is finding the proper files
PATH_NAME = dirname(dirname(abspath(__file__)))
devn = "deviations"
ilm = "ilm"
pro = "pro_numbers"
dx_ilm = "dx_ilm"

if isfile(PATH_NAME + '/assets/{}.xlsx'.format(devn)) and isfile(PATH_NAME + '/assets/{}.xls'.format(ilm)):
    print("Here")
else:
    print("Not Here")

sleep (20)
'''

'''
#This test is to make sure all libraries are imported
import traceback
import time

try:
    import os
    import csv
    import sys
    from itertools import chain
    from os.path import dirname, abspath, isfile
    import pandas as pd
    from xlsxwriter.workbook import Workbook
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

    print ("imports successful")
    time.sleep (20)
    
except:
    traceback.print_exc()
    time.sleep (20)
'''
