#!/usr/bin/env python
####################################################################################
#   Author:         Justin Chung, Daniel Brown
#
#   Last modifed:   By Justin: 2/26/2018   By Daniel Brown: /2021
#
#   Justin Comments:This file contains some scripting functions for fetching data
#                   from ETQ Reliance and from support files in /assets
#
#   Daniel Comments:I added an "canint" function. I have edited the "DownloadDHR" function.
#                   I will need to add a NC checker function to replace the old one. Change
#                   "get_ncr_numbers" name to "get_ncr_numbers_old". Added "get_nc_status" 
#                   and "open_Fiori_NC" functions.                
#
####################################################################################
import os
import csv
import sys
import time
from itertools import chain
from os.path import dirname, abspath, isfile
import pandas as pd
from xlsxwriter.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

#Daniel: I imported these libraries for the "get_nc_status" function
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

pathName = dirname(dirname(abspath(__file__)))

####################################################################################

def open_Fiori_NC(NCList):
    #Declare variables
    TabCount = 1

    # Get appropriate driver for current Google Chrome Version
    driver = webdriver.Chrome(ChromeDriverManager().install())
    
    #Check if NCList has any contents
    if not NCList:
        print("No NC's!")
        return None

    #Check the status of each NC
    for NCnum in NCList: 
        #Format NCList contents (generated by "edhr.parse_nc" function)
        NCnum = NCnum[-9:]
    
        #Change URL to open the Fiori NC directly
        NCurl = ('https://sapfioriprd.illumina.com/sap/bc/ui5_ui5/ui2/ushell/shells/abap/FioriLaunchpad.html?'
            'sap-client=100&sap-language=EN#ZP2D_NC_MANAGEQN_SM-manage&/QNInitiate/{}'.format(NCnum))

        #Open next NC in another tab
        if TabCount >= 2:
            driver.execute_script("window.open('{}');".format(NCurl))
            TabDiff = TabCount -1
            driver.switch_to.window(driver.window_handles[TabDiff])
            driver.get(NCurl)
            
        else:
            driver.get(NCurl)
        
        TabCount += 1
        
    
        #Wait for Fiori to open
        #Code from https://stackoverflow.com/questions/26086965/wait-until-loader-disappears-python-selenium/40903207
        SHORT_TIMEOUT  = 30   # give enough time for the loading element to appear
        LONG_TIMEOUT = 90  # give enough time for loading to finish
        LOADING_ELEMENT_XPATH = '//section[@class="sapMDialogSection"]/div/div'
        
        try:
        # wait for loading element to appear - required to prevent prematurely checking if element
        # has disappeared, before it has had a chance to appear
            WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

        # then wait for the element to disappear
            WebDriverWait(driver, LONG_TIMEOUT
                ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

        except TimeoutException:
            # if timeout exception was raised - it may be safe to assume loading has finished, however this may not 
            # always be the case, use with caution, otherwise handle appropriately.
            pass
        #End of stackoverflow code   
    input('Press Enter to close Fiori windows')
    try:
        for i in range(TabCount-1):
            driver.switch_to.window(driver.window_handles[0])
            driver.close()
    except:
        a=1

def get_nc_status(NCList):
    #Declare variables
    NCDic = {}
    TabCount = 1

    #Setup ChromeOptions and get appropriate chrome driver
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chromeOptions)
    
    #Check if NCList has any contents
    if not NCList:
        return NCDic
    
    #Check the status of each NC
    for NCnum in NCList: 
        #Format NCList contents (generated by "edhr.parse_nc" function)
        NCnum = NCnum[-9:]
    
        #Change URL to open the Fiori NC directly
        NCurl = ('https://sapfioriprd.illumina.com/sap/bc/ui5_ui5/ui2/ushell/shells/abap/FioriLaunchpad.html?'
            'sap-client=100&sap-language=EN#ZP2D_NC_MANAGEQN_SM-manage&/QNInitiate/{}'.format(NCnum))

        #Open next NC in another tab
        if TabCount >= 2:
            driver.execute_script("window.open('{}');".format(NCurl))
            TabDiff = TabCount -1
            driver.switch_to.window(driver.window_handles[TabDiff])
            driver.get(NCurl)
            
        else:
            driver.get(NCurl)
        
        TabCount += 1
        
    
        #Wait for Fiori to open
        #Code from https://stackoverflow.com/questions/26086965/wait-until-loader-disappears-python-selenium/40903207
        SHORT_TIMEOUT  = 30   # give enough time for the loading element to appear
        LONG_TIMEOUT = 90  # give enough time for loading to finish
        LOADING_ELEMENT_XPATH = '//section[@class="sapMDialogSection"]/div/div'
        
        try:
        # wait for loading element to appear - required to prevent prematurely checking if element
        # has disappeared, before it has had a chance to appear
            WebDriverWait(driver, SHORT_TIMEOUT
                ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

        # then wait for the element to disappear
            WebDriverWait(driver, LONG_TIMEOUT
                ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

        except TimeoutException:
            # if timeout exception was raised - it may be safe to assume loading has finished, however this may not 
            # always be the case, use with caution, otherwise handle appropriately.
            pass
        #End of stackoverflow code
        
        #Collect information about the NC
        InitiateButton = driver.find_element_by_xpath('//div[@class="sapMITBHead"]/div/div/span')
        InitiateButton.click()
        
        NCHeader = driver.find_element_by_xpath('//*[contains(text(),"Current Phase")]').text
        
        #Collect all items in the PIS table
        #PISrowItem is a string
        try:
            for PISnum in range(1,9):
                PISrowItem = driver.find_element_by_xpath('//a[text()="000{}"]/parent::td/following-sibling::td/span'.format(PISnum)).text
                if PISnum == 1:
                    NCDic[NCHeader] = [PISrowItem]
                else:
                    NCDic[NCHeader].append(PISrowItem)
            print('PIS has more than 9 items. Please manually check them.')
        except NoSuchElementException:
            a=1
        
        
        #As great as it would be for the script to pull the consolidated tasks, this is going to take
        #too much time to code. I should just open the webpages
        #Collect Tasks if PISrowItem indicates an NCI
        
        #if PISrowItem == '15033616':
        #    ActionsList = driver.find_element_by_xpath('//bdi[contains(text(),"Actions")]')
        #    ActionsList.click()
        #    print('action list deployed')
        #    sleep(10)
        #    ActionsListConsTasks = driver.find_element_by_xpath('//div[contains(text(),"Consolidated Tasks")]')
        #    ActionsListConsTasks.click()
        #    
        #   #xpath for consolidated tasks table: //tbody[@ id="ConsolidatedTasktable-tblBody"]

    #I may have to add this part to a different part of the program
    #os.system('cls')
    #for i in NCDic:
        #print(i," PIS Material(s):",NCDic[i])
    for i in range(TabCount-1):
        driver.switch_to.window(driver.window_handles[0])
        driver.close()
    return NCDic


def get_ncr_numbers_old(SN, username, password):
    ''' Get NCR numbers from ETQ Reliance '''
    try:
        # Open browser and navigate to ETQ Reliance
        chrome_options = Options()
        chrome_options.add_argument('--log-level=3')

        # Get appropriate driver for current Google Chrome Version
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

        url = "https://ussd-prd-etqr02.illumina.com/relianceprod/reliance?"
        "ETQ$CMD=CMD_OPEN_HOMEPAGE&ETQ$SCREEN_WIDTH=1280"
        driver.get(url)
        
        # Login
        inputUsername = driver.find_element_by_name('ETQ$LOGIN_USERNAME')
        inputPassword = driver.find_element_by_name('ETQ$LOGIN_PASSWORD')
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


def get_devn_numbers(filename):
    ''' Parse deviations Excel file and extract Eng. Record numbers '''
    try:
        path = pathName + '/assets/{}.xlsx'.format(filename)
        xl = pd.ExcelFile(path)
        df = xl.parse()
        return list(chain.from_iterable(df[['Eng. Record']].drop_duplicates().values.tolist()))
    except FileNotFoundError:
        print("Missing file {}.xlsx in {}".format(filename, pathName + '/assets/'))
        sys.exit(0)


def get_ilm_numbers(filename):
    ''' Parse the ILM Excel file and extract the ILM numbers '''
    try:
        path = pathName + '/assets/{}.xls'.format(filename)
        xl = pd.ExcelFile(path)
        df = xl.parse(skiprows=1)
        return list(chain.from_iterable(df[['Line #']].drop_duplicates().values.tolist()))
    except FileNotFoundError:
        print("Missing file {}.xls in {}".format(filename, pathName + '/assets/'))
        sys.exit(0)


def get_pro_numbers(filename):
    ''' Get PRO numbers and associated serial and material number '''
    nums = {}
    try:
        path = pathName + '/assets/{}.xlsx'.format(filename)
        xl = pd.ExcelFile(path)
        df = xl.parse()
        for row in df.itertuples():
            # { SN: (PRO, MN) }
            nums[row[3]] = (row[1], row[2])
        return nums

    except FileNotFoundError:
        print("Missing file {}.xlsx in {}".format(filename, pathName + '/assets/'))
        sys.exit(0)


def downloadDHR(SN):
    #Export and download eDHR file for given SN. Save that file in two formats (CSV, XLSX)
    #Daniel: I changed the file slashes from "/" to "\" in the cases of file paths
    eDHR_FilesPath = "{}\eDHR_Files".format(pathName)

    #Daniel: Made it possible to choose to update an existing eDHR file
    #ERROR: INCLUDING THE INPUT COMMAND THROWS AN INDENTATION ERROR ON THE IF STATEMENT FOLLOWING IT
    if isfile(eDHR_FilesPath + "\{}.csv".format(SN)):
        updateFile = input('\nWARNING: eDHR file already exists! Update file? Y/N: ')
        if updateFile.upper()[0] == 'Y':
            print('Replacing eDHR...')
            os.remove(eDHR_FilesPath + "\{}.csv".format(SN))
        else:
            print('eDHR replacement skipped\n')
            time.sleep(1)
            return

    # Configure Chrome profile to download to this script's directory
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : eDHR_FilesPath}
    chrome_options.add_experimental_option("prefs",prefs)
    chrome_options.add_argument('--log-level=3')
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)

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
    print ("eDHR Download Successful")


def csvToExcel(SN):
    ''' Open CSV and write contents to a new Excel Worksheet file with SN as the filename '''
    eDHR_FilesPath = "{}\eDHR_Files".format(pathName)

    workbook = Workbook("{}\{}.xlsx".format(eDHR_FilesPath, SN))
    worksheet = workbook.add_worksheet()
    
    #Daniel Brown added the "if" statement to prevent the program from terminating at this point. 
    if isfile("{}\eDHR - Instrument - Detail.csv".format(eDHR_FilesPath)):
        with open("{}\eDHR - Instrument - Detail.csv".format(eDHR_FilesPath), 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()


def get_dx_cal_dates(dx_filename):
    ''' Parse dx ilm file '''
    try:
        dx_cal_dates = dict()
        path = pathName + '/assets/{}.xls'.format(dx_filename)
        xl = pd.ExcelFile(path)
        df = xl.parse(usecols='A,I', skiprows=1)
        for row in df.iterrows():
            dx_cal_dates[row[1].values[0]] = row[1].values[1]
        return dx_cal_dates
    except FileNotFoundError:
        print("Missing file {}.xls in {}".format(dxFileName, pathName + '/Support/'))
        sys.exit(0)
 
#Daniel: this is the integer function checker. I added it since some of the 
#        code fails when certain variables cannot be converted to integers.

def canint(value):
        try:
            int(value)
            return True
        except ValueError:
            return False