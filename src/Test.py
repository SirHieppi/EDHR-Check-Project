####################################################
# By: Daniel Brown (dbrown5)
#
# Comment: This file is tests parts of the eDHR checker
####################################################

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

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

#new commands imported for a wait attempt
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import linecache
import traceback

import edhr
import support

#This code will pull statuses of NC's from Fiori. When I incorperate it, I will need an option
#not to exicute the code in case of any changes to our NC interface.
#Also, this code only works with MiSeq instruments

try:
    
    def WaitForFiori():
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
    
    
    def get_nc_status(NCList):
        #Declare variables
        NCDic = {}
        TabCount = 1

        #Check if NCList has any contents
        if not NCList:
            print("No NC's")
            return NCDic
    
        #Setup Chromedriver
        chromedriver = r"\\ushw-file\Users\transfer\edhr_checker\eDHR\eDHR\src\chromedriver_win32\chromedriver.exe" 
        chromeOptions = webdriver.ChromeOptions()
        driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)
        
        #Check the status of each NC
        for NCnum in NClist: 
            NCnum = str(NCnum)
        
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
            
            WaitForFiori()
            
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
                print("")
            
            
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
        os.system('cls')
        for i in NCDic:
            print(i," PIS Material(s):",NCDic[i])
        input('Press Enter to close windows.')
        for i in range(TabCount-1):
            driver.switch_to.window(driver.window_handles[0])
            driver.close()
        return NCDic
    
    #TEST WHAT HAPPENS WITH AN EMPTY LIST
    NCRs = []
    get_nc_status(NCRs)
    
except Exception:
    traceback.print_exc()
    sleep (60)    