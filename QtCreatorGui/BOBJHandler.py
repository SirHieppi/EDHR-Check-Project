from win32com.client.gencache import __init__
import win32com.client
import win32api
import os
import json
from os import path

# Excel sheet names
productSummarySheet = os.getcwd() + "\\Section-1 Product Summary.csv"
deviationsSheet = os.getcwd() + "\\Section-3 Deviations.csv"
nonConformancesSheet = os.getcwd() + "\\Section-4 Non-Conformances.csv"
attachmentsSheet = os.getcwd() + "\\Section-5 Attachments.csv"
transactionDetailsSheet = os.getcwd() + "\\Section-7 Transaction Details.csv"


class BOBJHandler:
    def __init__(self):
        pass

    def extractProductSummary(self):
        pass

    def extractDeviations(self):
        pass

    def extractNonConformances(self):
        pass

    def extractAttachments(self):
        pass

    def extractTransactionDetails(self):
        pass

    def checkAttachments(self):
        pass



