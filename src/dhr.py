#!/usr/bin/env python
####################################################################################
#   Author:         Justin Chung, Daniel Brown
##
#   Last modifed:   By Justin: 03/19/2019   By Daniel: 7/21/2021
#
#   Justin Comments:DeviceHistoryRecord is a Class used to store information about an
#					instrument that can be found in its' DHR.
#
#   Daniel Comments:I have made edits to the "_parse_deviations" function and added the
#                   "__iter__" function.
#
#	TODO:			Each platform/mn should have it's own parse_mes_transaction since
#					a lot of the stuff is hard-coded
#
####################################################################################
import edhr
import support
from os.path import dirname, abspath
import csv
import datetime

try:
    class DeviceHistoryRecord:

        PLATFORM_TRANSACTIONS_MAP = {
            '15033616': edhr.parse_mes_transactions_ruo,
            '15048976': edhr.parse_mes_transactions_ruo,
            '20013740': edhr.parse_mes_transactions_nova,
            '20014737': edhr.parse_mes_transactions_dx
        }

        def __init__(self, material_num=None, serial_num=None, xl=None, csv=None, write=False):
            self.material_num = material_num
            self.serial_num = serial_num
            self.pro_num = None
            self.mnlist = [15033616, 15048976, 20014737, 20014054, 20013740]
            
            self.deviations = None
            self.nc = None
            self.mt = None

            self.xl = xl
            self.csv = csv

            self.row_offsets = edhr.parse_csv(self.csv)

            self.errors = []
            self.bad_transactions = None
            self.fio_data_points = None
            self.comments = None

            self.write = write
            self.log = None
            #Daniel: Added this dictionary
            self.FioriNCs = {}

            
        def _parse_product_summary(self, pro_numbers):
            details = edhr.parse_product_summary(self.xl)

            try:
                self.pro_num = details['PRO']
                if not details['SN']:
                    self.errors.append(ValueError('Serial number {} is not consistent'.format(self.serial_num)))
                if details['PRO'] != pro_numbers[self.serial_num][0]:
                    self.errors.append(ValueError('PRO {} is not consistent'.format(details['PRO'])))
                if str(details['MN']) != str(self.material_num) or str(details['MN']) != str(pro_numbers[self.serial_num][1]):
                    self.errors.append(ValueError('Material number {} is not consistent'.format(details['MN'])))
            except KeyError:
                self.errors.append(ValueError('Serial number does not exist. Please check that all files in /assets are up-to-date.'))

            # Product summary
            print('   ---------------------------------------------')
            print('   | SN: {} | PRO: {} | MN: {} |'.format(self.serial_num, self.pro_num, self.material_num))
            print('   ---------------------------------------------')

        #Daniel: I added the "canint" check
        def _parse_deviations(self, devn_numbers):
            self.deviations = edhr.parse_deviations(self.xl, self.row_offsets)
            for op, deviations in self.deviations.items():
                for deviation in deviations:
                    if type(deviation) is int and deviation not in devn_numbers:
                        self.errors.append(ValueError('{}: deviation {} not found in active deviations'.format(op, deviation)))
                    elif type(deviation) is str and 'REMOVE' not in str(deviation) and support.canint(deviation) is True:
                        if int(deviation) not in devn_numbers:
                            self.errors.append(ValueError('{}: deviation {} not found in active deviations'.format(op, deviation)))


        #Daniel: I can use this to check the Fiori NC's
        def _parse_nc(self, ncr_numbers):
            self.nc = edhr.parse_nc(self.xl, self.row_offsets)

            #Daniel: I excluded these since they are no longer relevant
            #for key in self.nc:
                #if key not in ncr_numbers:
                    #if "REMOVE" in key:
                        #pass
                    #elif key + "-REMOVED" not in self.nc:
                        #print('here', key) 
                        #self.errors.append(ValueError('Incorrect NC: {} not found in ETQ'.format(key)))
            for ncr in ncr_numbers.keys():
                if ncr_numbers[ncr][0] == self.material_num and ncr not in self.nc:
                    self.errors.append(ValueError("Missing NC: {} found in ETQ, but not recorded in DHR".format(ncr)))
                if ncr_numbers[ncr][0] == self.material_num and ncr_numbers[ncr][1] not in ["Verify", "Complete"]:
                    self.errors.append(ValueError("{} is a NCI and is still in phase {}. Must be in at least Verify in order to ship".format(ncr, ncr_numbers[ncr][1])))
                if ncr_numbers[ncr][0] != self.material_num and ncr_numbers[ncr][1] in ["Initiate", "Event Determination"]:
                    self.errors.append(ValueError("{} is a NCM and is still in phase {}. Must be in at least Evaluate in order to ship".format(ncr, ncr_numbers[ncr][1])))
            

        def _parse_mes_transactions(self, ilm_numbers, dx_ilm_numbers=None):
            self.bad_transactions, self.fio_data_points, self.comments = self.PLATFORM_TRANSACTIONS_MAP[str(self.mnlist[self.material_num - 1])](self.xl, self.row_offsets, dx_ilm_numbers if self.material_num == '20014737' else ilm_numbers)
            for transaction in self.bad_transactions:
                self.errors.append(ValueError(transaction))


        def _print(self):
            if self.write:
                save_path = '//ushw-file/Users/transfer/edhr_checker/logs/{} {}.csv'.format(self.serial_num, str(datetime.datetime.now().strftime('%Y-%m-%d__%H_%M_%S')))
                
                with open(save_path, mode='a', newline='', encoding='utf-8') as csv_file:
                    writer = csv.writer(csv_file, delimiter=',')

                    # Deviations
                    print('Deviations:')
                    print('   ------------------------------------------------------------')
                    print('   |{:<12}|{:<45}|'.format('Operation', 'Deviations'))
                    print('   ------------------------------------------------------------')

                    fieldnames = ['Operation', 'Deviations']
                    writer.writerow(fieldnames)
                    for op, deviation in self.deviations.items():
                        string = ", ".join(map(str, deviation))
                        print('   |{:<12}|{:<45}|'.format(op, string))
                        writer.writerow([op, string])
                
                    print('   ------------------------------------------------------------')
                    writer.writerow([])

                    # Non-conformances
                    print('Non-conformances:')

                    fieldnames = ['NCR#']
                    writer.writerow(fieldnames)
                    #Daniel: Replaced 'self.nc' with 'self.FioriNCs'
                    if len(self.FioriNCs) > 0:
                        for nc in self.FioriNCs:
                            print('   {}'.format(nc),' PIS Material(s):',self.FioriNCs[nc])
                    else:
                        print('   None')
                    writer.writerow([])

                    # FIO data points
                    print('FIO Data Points:')
                    print('   --------------------------------------------------------------------------------------------------------')
                    print('   |{:<29}|{:<30}|{:<16}|{:<24}|'.format('Step', 'Task', 'Result', 'Timestamp'))
                    print('   --------------------------------------------------------------------------------------------------------')

                    writer.writerow('FIO Data Points')
                    fieldnames = ['Step', 'Task', 'Result', 'Timestamp']
                    writer.writerow(fieldnames)
                    for point in self.fio_data_points:
                        print('   |{:<29}|{:<30}|{:<16}|{:<24}|'.format(point[0], point[1], point[2], point[3]))
                        writer.writerow([point[0], point[1], point[2], point[3]])
                    print('   --------------------------------------------------------------------------------------------------------')
                    writer.writerow([])

                    # Comments
                    print('Comments: ')
                    print('   -------------------------------------------------------------------------------------------------------')
                    print('   |{:<22}|{:<24}|{}'.format('Task', 'Timestamp', 'Comment'))
                    print('   -------------------------------------------------------------------------------------------------------')

                    writer.writerow('Comments')
                    fieldnames = ['Task', 'Timestamp' ,'Comment']
                    writer.writerow(fieldnames)
                    for comment in self.comments:
                        print('   |{:<22}|{:<24}|{}'.format(comment[0], comment[1], comment[2]))
                        writer.writerow([comment[0], comment[1], comment[2]])
                    print('   -------------------------------------------------------------------------------------------------------')
                    writer.writerow([])

                    # Errors
                    print('Check these transactions for potential errors')
                    writer.writerow(['Check these for potential errors'])
                    if len(self.errors) > 0:
                        for error in self.errors:
                            print('   - {}'.format(error))
                            writer.writerow([error])
                    else:
                        print('   None')
                    writer.writerow([])
                    
            else:
                # Deviations
                print('Deviations:')
                print('   ----------------------------------------------------------')
                print('   |{:<10}|{:<45}|'.format('Operation', 'Deviations'))
                print('   ----------------------------------------------------------')
                if len(self.deviations) > 0:
                    for op, deviation in self.deviations.items():
                        string = ", ".join(map(str, deviation))
                        print('   |{:<10}|{:<45}|'.format(op, string))
                else:
                    print('   None')
                print('   ----------------------------------------------------------')

                # Non-conformances
                print('Non-conformances:')
                #Daniel: Replaced 'self.nc' with 'self.FioriNCs'
                if len(self.FioriNCs) > 0:
                    for nc in self.FioriNCs:
                        print('   {}'.format(nc),' PIS Material(s):',self.FioriNCs[nc])
                else:
                    print('   None')

                # FIO data points
                print('FIO Data Points:')
                print('   -------------------------------------------------------------------------------------------------')
                print('   |{:<22}|{:<30}|{:<16}|{:<24}|'.format('Step', 'Task', 'Result', 'Timestamp'))
                print('   -------------------------------------------------------------------------------------------------')
                for point in self.fio_data_points:
                    print('   {}'.format(point))
                print('   -------------------------------------------------------------------------------------------------')

                # Comments
                print('Comments: ')
                print('   -------------------------------------------------------------------------------------------------------')
                print('   |{:<22}|{:<24}|{}'.format('Task', 'Timestamp', 'Comment'))
                print('   -------------------------------------------------------------------------------------------------------')
                for comment in self.comments:
                    print('   {}'.format(comment))
                print('   -------------------------------------------------------------------------------------------------------')

                # Errors
                print('Check these transactions for potential errors: ')
                if len(self.errors) > 0:
                    for error in self.errors:
                        print('   - {}'.format(error))
                else:
                    print('   None')

        def check(self, ilm_numbers, devn_numbers, pro_numbers, ncr_numbers, dx_ilm_numbers):
            self._parse_product_summary(pro_numbers)
            self._parse_deviations(devn_numbers)
            self._parse_nc(ncr_numbers)
            #Daniel: Start incorperating the "get_nc_status" function
            #GOAL: Have this run before the results print.
            self.FioriNCs = support.get_nc_status(self.nc)
            
            #Daniel: End Code

            self._parse_mes_transactions(ilm_numbers=ilm_numbers, dx_ilm_numbers=dx_ilm_numbers)

            self._print()
            
            #Daniel: Adding this to open Fiori windows for the user's convience
            if self.FioriNCs:
                OpenFioriNCs = input("Do you want to open NC's on this instrument? Y/N: ")
                if OpenFioriNCs.upper() == "Y":
                    support.open_Fiori_NC(self.nc)

            
        #Daniel: this code makes the DeviceHistoryRecord class iterable  
        def __iter__(self):
            return iter(range(3))
        
        #Daniel: End code
            

except Exception as e:
    print ("Error!")
    print (e)
    sleep (60)