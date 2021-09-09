#!/usr/bin/env python
#####################################################################################
#   Author:         Justin Chung
#
#   Last modifed:   03/19/2019
#
#   Comments:       The purpose of this script is to automate the process of checking
#                   an instrument's eDHR. Checking a DHR by hand is very time consuming,
#                   inefficent, and error-prone. This script is designed to assist this
#                   task only.
#
#   TODO:           - Come up with a better UI
#                   - Do more error handling for user input
#
#####################################################################################
import support
import dhr
import pandas as pd
import sys
import getpass
from os.path import dirname, abspath, isfile

def main():
    PATH_NAME = dirname(dirname(abspath(__file__)))
    devn_numbers = None
    ilm_numbers = None

    print('------------------------------------------------------------------------------------------------')
    print('|                         Welcome to the MiSeq RUO/FGx eDHR Checker                            |')
    print('------------------------------------------------------------------------------------------------')

    username = user_input("Illumina username: ")
    password = getpass.getpass("Illumina password: ")

    while True:
        devn = "deviations"
        ilm = "ilm"
        pro = "pro_numbers"
        dx_ilm = "dx_ilm"

        if isfile(PATH_NAME + '/assets/{}.xlsx'.format(devn)) and isfile(PATH_NAME + '/assets/{}.xls'.format(ilm)):
            devn_numbers = support.get_devn_numbers(devn)                     
            ilm_numbers = support.get_ilm_numbers(ilm)
            pro_numbers = support.get_pro_numbers(pro)
            dx_ilm_numbers = support.get_dx_cal_dates(dx_ilm)
            break
        else:
            print("\n\tWARNING: One or more Support files could not be found")

    while True:
        print("------------------------------------------------------------------------------------------------")
        while True:
            mn = user_input("Instrument material number: ").strip()
            sn = user_input("Instrument serial number: ").upper().strip()
            #                 RUO         FGx         DX                    NovaSeq
            if mn not in ['15033616', '15048976', '20014737', '20014054', '20013740']:
                print("Could not recognize platform with material number {}".format(mn))
            else:
                break
        print("Please wait while we download some files. . .")
        support.downloadDHR(sn)
        ncr_numbers = support.get_ncr_numbers(sn, username, password)
        xl = pd.ExcelFile(PATH_NAME + "/eDHR_Files/{}.xlsx".format(sn))
        csv = open(PATH_NAME + '/eDHR_Files/{}.csv'.format(sn), mode='r', encoding='utf8')
        print("------------------------------------------------------------------------------------------------")
        
        instrument = dhr.DeviceHistoryRecord(material_num=mn, serial_num=sn, xl=xl, csv=csv, write=True)
        instrument.check(ilm_numbers=ilm_numbers, devn_numbers=devn_numbers, pro_numbers=pro_numbers, ncr_numbers=ncr_numbers, dx_ilm_numbers=dx_ilm_numbers)

def user_input(msg):
    s = input(msg)
    if s.upper() == "Q" or s.upper() == "QUIT":
        sys.exit(0)
    return s

def valid_sn(sn):
    try:
        int(sn[1:])
        return True
    except:
        return False

if __name__ == "__main__":
    main()