"""
Notepad Automation Project

The project contain two major module:

1.Notepad: Contain all the functional test cases for notepad automation.
2.UpdateExcel: Contains all the method to update the result sheet.

Below libraries need to be installed:
pywinauto
pandas
xlrd
xlwt

"""

# Importing required modules
from notepad import *

n1 = Notepad()

# Calling Test Cases
n1.test_case_multilizer()


# Printing the test cases statistics
print("Total test case executed: ", excel_update.testcase_total)
print("Total test case passed: ", excel_update.testcase_passed)
print("Total test case failed: ", excel_update.testcase_failed)
