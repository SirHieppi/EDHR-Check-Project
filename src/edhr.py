#!/usr/bin/env python
####################################################################################
#   Author:         Justin Chung, Daniel Brown
#
#   Last modifed:   By Justin: 02/26/2018   By Daniel: 7/21/2021
#
#   Justin Comment: This file contains functions for parsing the DHR. DHR's all have
#                   the same format. The only difference is in MES Transacitions.
#                   This makes it necessary to create a new function for separate
#                   platforms.
#
#   Daniel Comment: Change in the "parse_product_summary" function. Change in the
#                   "parse_deviations" function
#
####################################################################################
import re
import math
from os.path import dirname, abspath
import pandas as pd
import numpy as np

#xlFile usually refers to the eDHR excel file

try:
    def parse_csv(csv):
        row_offsets = []

        with csv as file:
            count = 0
            for line in file.readlines():
                count += 1
                line = line.strip(',').strip()
                if not line:
                    row_offsets.append(count)
        return row_offsets
        

    def parse_product_summary(xlFile):
        productDetails = None

        df = xlFile.parse(nrows=1, usecols="A:D")
        
        '''
        # Daniel: this block of code produced an error, so I replaced it with the code
        # found between this comment and the "return productSummary" command
        for row in df.iterrows():
            productDetails = list(row[0])
            productDetails.append(row[1].values.tolist()[0])
            
        productSummary = dict(zip(['SN', 'MN', 'Type', 'PRO'], productDetails))
        '''
        
        df_list = df.values.tolist()
        productDetails = []
        
        for row in df_list[0]:
            productDetails.append(row)
            
        productSummary = dict(zip(['SN', 'MN', 'Type', 'PRO'], productDetails))
        
        return productSummary
        
   
    def parse_deviations(xlFile, row_offsets):
        devns = dict()
        
        df = xlFile.parse(skiprows=row_offsets[0], nrows=row_offsets[1] - row_offsets[0] - 2, usecols="A, C")
        
        for row in df.iterrows():
            op = str(row[1].values[0]).split('-')[0]
            deviation = str(row[1].values[1])
            if type(row[1].values[1]) is float and deviation != 'nan':
                deviation = str(row[1].values[1])[:-2]

            if deviation != 'nan':
                if op in devns:
                    devns[op].append(deviation)
                else:
                    devns[op] = [deviation]
        
        #Daniel: From this line to the return, I added this code
        #ERROR: When working on M07680, I get an error on line 79 in the dhr script
        #Script goes through each line in the eDHR looking for deviations listed
        df = xlFile.parse(skiprows=row_offsets[4], nrows=row_offsets[5] - row_offsets[4] - 2, usecols="A,C,G")
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

        #Delete any duplicate values within the 'devns' dictionary
        for op in devns:
            devns[op] = list(set(devns[op]))            
        return devns


    def parse_nc(xlFile, row_offsets):
        ncs = []
        
        df = xlFile.parse(skiprows=row_offsets[1], nrows=row_offsets[2] - row_offsets[1] - 2, usecols="C")
        
        for row in df.iterrows():
            ncrNumber = row[1].values[0]
            if (isinstance(ncrNumber, (int, np.integer, float, np.float)) and not math.isnan(ncrNumber)) or type(ncrNumber) is str:
                ncs.append("NCR-" + str(ncrNumber))
        return ncs


    def parse_attachments(xlFile, row_offsets):
        ''' Parse "Attachments" table of eDHR
        '''
        attachments = []
        
        df = xlFile.parse(skiprows=row_offsets[2], nrows=row_offsets[3] - row_offsets[2] - 2, usecols="A")
        
        for row in df.iterrows():
            attachments.append(row[1].values[0])
        return attachments


    def parse_mes_comments(xlFile, row_offsets):
        ''' Parse "Exception MES Transaction Details" of eDHR
        ''' 
        comments = dict()

        df = xlFile.parse(skiprows=row_offsets[3], nrows=row_offsets[4] - row_offsets[3] - 2, usecols="A,H")
        
        count = 0
        for row in df.iterrows():
            count += 1
            vals = row[1].values
            comments["{}.".format(count)] = str(vals[0]) + ' --> ' + str(vals[1])
        return comments


    def parse_mes_transactions_ruo(xlFile, row_offsets, ilmNumbers):
        ''' Parse "MES Transaction Details" table of the eDHR
            Check:
                (1) Bad ILM numbers by checking against spreadsheet
                (2) Perform and verify Burn-In are at least 9 hours apart
                (3) Find result values that are "Other" or "Not Performed", indicating a possible operator error
        '''
        df = xlFile.parse(skiprows=row_offsets[4], nrows=row_offsets[5] - row_offsets[4] - 2, usecols="A,C,D,E,F,G,K,L,M")

        bad_transactions = []
        fio_data_points = []
        comments = []
        time = None
        rework = False
        go_to_nc_path = False
        prev_op = None
        
        # Cols |   A   |   C   |      D      |    E    |    F    |    G    |        K        |      L      |    M    |
        keys = ['Step', 'Task', 'Description', 'Upper', 'Lower', 'Result', 'Transaction Type', 'Timestamp', 'Comment']
        for row in df.iterrows():
            temp = dict(zip(keys, row[1].values))
            
            # Check rework path
            if temp['Step'] == "NC/DEV/RW" and temp['Description'] == "NCR Number" and not rework:
                print("Rework for NCR-{}, which failed in {}".format(temp['Result'], prev_op))
                print('   ------------------------------------------------------------------------------------')
                print('   |{:<32}|{:<32}|{:<16}|'.format('Task', 'Description', 'Result'))
                print('   ------------------------------------------------------------------------------------')
                rework = True
            if rework:
                print("   |{:<32}|{:<32}|{:<16}|".format(temp['Task'], temp['Description'], temp['Result']))
                if (temp['Step'] == "Conf NC/DEV/RW-Exit" and temp['Result'] == "No") or temp['Transaction Type'] == "ilmMoveNonStdCurrentWorkflow":
                    rework = False
                    print('   ------------------------------------------------------------------------------------')

            # Go to NC path before executing main task
            if go_to_nc_path:
                if temp['Task'] != "Record NC/Dev and/or Rework":
                    bad_transactions.append('[{}]: {}').format(temp['Timestamp'], "Go to NC Path before Executing Task")
                    go_to_nc_path = False

            if temp['Task'] == 'Go to NC/DEV/RWPath' and temp['Result'] == 'Go to NC Path' and temp['Transaction Type'] == 'ExecuteTask':
                print('Go to NC path BEFORE executing task')
                go_to_nc_path = True

            # FIO data points
            if type(temp['Upper']) is float and type(temp['Lower']) is float and math.isnan(temp['Upper']) and math.isnan(temp['Lower']) and type(temp['Result']) is str:
                if re.search(r"(\(°*\s*\w+\s*\))", temp['Result']):
                    #fio_data_points.append('|{:<22}|{:<30}|{:<16}|{:<24}|'.format(temp['Task'], temp['Description'], temp['Result'], temp['Timestamp']))
                    fio_data_points.append((temp['Task'], temp['Description'], temp['Result'], temp['Timestamp']))

            # Comments
            if type(temp['Comment']) is str:
                #comments.append("|{:<22}|{:<24}|{}".format(temp['Step'], temp['Timestamp'], temp['Comment']))
                comments.append((temp['Step'], temp['Timestamp'], temp['Comment']))

            # Check ILM numbers
            if type(temp['Description']) is str and ('Item#' in temp['Description'] or 'Item Number' in temp['Description']) and temp['Result'] not in ilmNumbers:
                bad_transactions.append('[{}]: {} ({}) --> ({})'.format(temp['Timestamp'], "Inconsistent ILM #", temp['Description'], temp['Result']))

            # Result is "Other"
            if temp['Result'] == "Other":
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], 'Result is "Other"', temp['Description']))

            # Result is "Not Performed"
            if temp['Result'] == "Not Performed":
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], 'Result is "Not Performed"', temp['Description']))

            # Check that perform and verify burn-in are at least 19 hours apart
            if time and temp['Description'] == "Bubble Test" and time_difference(time, pd.to_datetime(temp['Timestamp'])) < pd.Timedelta(8, unit='h'):
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], "Verify BurnIn executed prematurely", temp['Description']))

            # Check that FIT run is not executed prematurely
            if time and temp['Task'] == "Verif 2x301 Cycl Compltd" and time_difference(time, pd.to_datetime(temp['Timestamp'])) < pd.Timedelta(24, unit='h'):
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], "Verify FIT executed prematurely", temp['Description']))

            # Check that MCS VCL is at least 60 minutes
            if time and temp['Description'] == "MCS VCL Timestamp" and time_difference(time, pd.to_datetime(temp['Timestamp'])) < pd.Timedelta(0.5, unit='h'):
                bad_transactions.append('[{}]: {} ({})}'.format(temp['Timestamp'], "MCS VCL executed prematurely", temp['Description']))

            time = pd.to_datetime(temp['Timestamp'])
            prev_op = temp['Step']

        return (bad_transactions, fio_data_points, comments)

    def parse_mes_transactions_dx(xlFile, row_offsets, dxIlmNumbers):
        df = xlFile.parse(skiprows=row_offsets[4], nrows=row_offsets[5] - row_offsets[4] - 2, usecols="A,C,D,E,F,G,K,L,M")

        bad_transactions = []
        fio_data_points = []
        comments = []
        time = None
        rework = False
        go_to_nc_path = False
        prev_op = None
        next_is_cal_date = False
        ilm_numbers = {}

         # Cols |   A   |   C   |      D      |    E    |    F    |    G    |        K        |      L      |    M    |
        keys = ['Step', 'Task', 'Description', 'Upper', 'Lower', 'Result', 'Transaction Type', 'Timestamp', 'Comment']
        for row in df.iterrows():
            temp = dict(zip(keys, row[1].values))

             # Check rework path
            if temp['Step'] == "NC/DEV/RW" and temp['Description'] == "NCR Number" and not rework:
                print("Rework for NCR-{}, which failed in {}".format(temp['Result'], prev_op))
                print('   ------------------------------------------------------------------------------------')
                print('   |{:<32}|{:<32}|{:<16}|'.format('Task', 'Description', 'Result'))
                print('   ------------------------------------------------------------------------------------')
                rework = True
            if rework:
                print("   |{:<32}|{:<32}|{:<16}|".format(temp['Task'], temp['Description'], temp['Result']))
                if (temp['Step'] == "Conf NC/DEV/RW-Exit" and temp['Result'] == "No") or temp['Transaction Type'] == "ilmMoveNonStdCurrentWorkflow":
                    rework = False
                    print('   ------------------------------------------------------------------------------------')

            # Go to NC path before executing main task
            if go_to_nc_path:
                if temp['Task'] != "Record NC/Dev and/or Rework":
                    bad_transactions.append('[{}]: {}').format(temp['Timestamp'], "Go to NC Path before Executing Task")
                    go_to_nc_path = False

            if temp['Task'] == 'Go to NC/DEV/RWPath' and temp['Result'] == 'Go to NC Path' and temp['Transaction Type'] == 'ExecuteTask':
                print('Go to NC path BEFORE executing task')
                go_to_nc_path = True

            # FIO data points
            if type(temp['Upper']) is float and type(temp['Lower']) is float and math.isnan(temp['Upper']) and math.isnan(temp['Lower']) and type(temp['Result']) is str:
                if re.search(r"(\(°*\s*\w+\s*\))", temp['Result']):
                    fio_data_points.append((temp['Task'], temp['Description'], temp['Result'], temp['Timestamp']))

            # Comments
            if type(temp['Comment']) is str:
                comments.append((temp['Step'], temp['Timestamp'], temp['Comment']))


            # Result is "Other"
            if temp['Result'] == "Other":
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], 'Result is "Other"', temp['Description']))

            # Result is "Not Performed"
            if temp['Result'] == "Not Performed":
                bad_transactions.append('[{}]: {} ({})'.format(temp['Timestamp'], 'Result is "Not Performed"', temp['Description']))

            # Check ILM numbers exists
            if type(temp['Description']) is str and ('Item#' in temp['Description'] or 'Item Number' in temp['Description']) and temp['Result'] not in dxIlmNumbers:
                bad_transactions.append('[{}]: {} ({}) --> ({})'.format(temp['Timestamp'], "Inconsistent ILM #", temp['Description'], temp['Result']))
                next_is_cal_date = True

            # Check due dates are not expired
            if  type(temp['Description']) is str and ('CalDue' in temp['Description'] or 'Cal Due' in temp['Description']) and time_difference(pd.to_datetime(temp['Result']), pd.to_datetime(temp['Timestamp'])) > pd.Timedelta(0, unit='h'):
                bad_transactions.append('{:<20} at {:<30} {:<20}'.format(temp['Description'], temp['Timestamp'], "Cal due date is expired"))
            
            time = pd.to_datetime(temp['Timestamp'])
            prev_op = temp['Step']

        return (bad_transactions, fio_data_points, comments)


    def parse_mes_transactions_nova(xlFile, row_offsets, ilmNumbers):
        df = xlFile.parse(skiprows=row_offsets[4], nrows=row_offsets[5] - row_offsets[4] - 2, usecols="A,C,D,E,F,G,K,L,M")

        bad_transactions = []
        fio_data_points = []
        comments = []
        time = None
        rework = False

        # Cols |   A   |   C   |      D      |    E    |    F    |    G    |        K        |      L      |    M    |
        keys = ['Step', 'Task', 'Description', 'Upper', 'Lower', 'Result', 'Transaction Type', 'Timestamp', 'Comment']
        for row in df.iterrows():
            temp = dict(zip(keys, row[1].values))

            # Check ILM numbers exists
            if type(temp['Description']) is str and 'ILM#' in temp['Description'] and temp['Result'] not in ilmNumbers:
                bad_transactions.append('[{}]: {} ({}) --> ({})'.format(temp['Timestamp'], "Inconsistent ILM #", temp['Description'], temp['Result']))

            time = pd.to_datetime(temp['Timestamp'])

        return (bad_transactions, fio_data_points, comments)


    def time_difference(before, after):
        ''' Return the difference in time between after and before as a pd.Timestamp object
        '''
        return after - before


    def is_date(st):
        try:
            dparser.parse(st)
            return True
        except ValueError:
            return False

    def parseDxIlmNumbers(xlFile, row_offsets, ilmNumbers, dx_cal_dates):
        df = xlFile.parse(skiprows=rowOffsets[4], nrows=rowOffsets[5] - rowOffsets[4] - 2, usecols="A,B,C,D,E,F,G,K,L,M")
        col_temp = df['DataValue']
        two_col = df[['DataValue', 'Textbox112']]

        col_temp = [x for x in col_temp if str(x) != 'nan']

        col_temp = pd.DataFrame(data=col_temp)

        count = 0
        matching = []
        # matching = [s for s in col_temp if 'ILM' in s or '2019' in s]
        keys = ['Result', 'UDC']

        for s in two_col.iterrows():
            count += 1
            rows = len(two_col)
            temp = dict(zip(keys, s[1].values))
            print(temp)

            if count > 1 and count < rows:
                previous = two_col.iloc[count-2]
                previous = dict(zip(keys, previous))

                next_item = two_col.iloc[count]
                next_item = dict(zip(keys, next_item))

                # Gets a specific column
                # fr = two_col.iloc[:, 0]
                udc_current = str(temp.get('UDC'))
                udc_prev = str(previous.get('UDC'))
                udc_nxt = str(next_item.get('UDC'))

                current = str(temp.get('Result'))
                prev = str(previous.get('Result'))
                nxt = str(next_item.get('Result'))

                if 'ILM' in current and (is_date(prev)==True or is_date(nxt)==True):

                    # Check if the value next to ilm is date
                    if is_date(nxt) == True and is_date(prev)== False:
                        matching.append(current)
                        matching.append(nxt)

                    # Check if previous value is date
                    if is_date(nxt) == False and is_date(prev) == True:
                        matching.append(current)
                        matching.append(prev)

                    # Check if Middle section for dates
                    if is_date(prev) == True and is_date(nxt) == True:
                        dot = '.'
                        remove_nxt = udc_nxt.split(dot, 1)[1]
                        remove_current = udc_current.split(dot, 1)[1]

                        pound = '#'
                        hypen = '-'

                        # if remove_current == remove_nxt:
                        #     matching.append(current)
                        #     matching.append(nxt)
        return matching

except Exception as e:
    print ("\nError!")
    print (e)
    sleep (60)