####################################################################################
#
# By: Daniel Brown (dbrown5)
#
# Last Edited:
#
# Comment: Make sure to resolve issue indicated by the "CHANGE NEED" sections
#
####################################################################################

#Need to test this code


try:
    #Import Libraries
    import traceback
    
    import os
    import csv
    import sys
    import time
    from time import sleep
    from itertools import chain
    from os.path import dirname, abspath, isfile
    import pandas as pd
    from xlsxwriter.workbook import Workbook
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
    import support
    import dhr
    import getpass


    if __name__ == "__main__":
    
        PATH_NAME = dirname(dirname(abspath(__file__)))
        devn_numbers = None
        ilm_numbers = None
        
        #Setting ncr_numbers to an empty list since the NCR script does not work anymore
        ncr_numbers={}
        
        print('------------------------------------------------------------------------------------------------')
        print('|                         Welcome to the MiSeq RUO/FGx eDHR Checker                            |')
        print('------------------------------------------------------------------------------------------------')
        
        #Asking user for instrument information
        print('\n                       Short Inst Material Number')
        print('--------1--------   ------2-----   -----3----- ---4----   --------5-------')
        print('MiSeqRUO:15033616   FGx:15048976   DX:20014737,20014054   NovaSeq:20013740')
        # mn = input("\nShort Instrument material number: ")
        SN = input("Instrument serial number: ")

        #Convert short Inst Material number to full number
        mnlist = [15033616, 15048976, 20014737, 20014054, 20013740]
        # # if int(mn) <= 6:
            # # mn = str(mnlist[int(mn)-1])
            
        mn = 5
    
        devn = "deviations"
        ilm = "ilm"
        pro = "pro_numbers"
        dx_ilm = "dx_ilm"

        #Get information from Excel sheets in the assets folder
        if isfile(PATH_NAME + r'\assets\{}.xlsx'.format(devn)) and isfile(PATH_NAME + r'\assets\{}.xls'.format(ilm)):
            devn_numbers = support.get_devn_numbers(devn)                     
            ilm_numbers = support.get_ilm_numbers(ilm)
            pro_numbers = support.get_pro_numbers(pro)
            dx_ilm_numbers = support.get_dx_cal_dates(dx_ilm)
        else:
            print("\n\tWARNING: One or more Support files could not be found!!!!!!")
            
        # print(str(devn_numbers) + "; " + str(ilm_numbers) + "; " + str(pro_numbers) + "; " + str(dx_ilm_numbers))
        
        #Download an Excel version of the eDHR and convert it to a CSV file
        support.downloadDHR(SN)
        support.csvToExcel(SN)
        
        xl = pd.ExcelFile(PATH_NAME + "/eDHR_Files/{}.xlsx".format(SN))
        csv = open(PATH_NAME + '\eDHR_Files\{}.csv'.format(SN), mode='r', encoding='utf8')
                
        #Run device checks
        instrument = dhr.DeviceHistoryRecord(material_num=mn, serial_num=SN, xl=xl, csv=csv, write=True)
        instrument.check(ilm_numbers=ilm_numbers, devn_numbers=devn_numbers, pro_numbers=pro_numbers, 
            ncr_numbers=ncr_numbers, dx_ilm_numbers=dx_ilm_numbers)
        
        print('\nCheck Completed.')
        input('Press Enter to close window.')

except Exception:
    traceback.print_exc()
    sleep (60)
    

