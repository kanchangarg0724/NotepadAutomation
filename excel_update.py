# Importing required libraries
import pandas as pd

testcase_total = 7
testcase_passed = 0
testcase_failed = 7


# Update the excel sheet on the basis of test execution
class UpdateExcel:

    def testcase_pass(self, row=0, col=0):
        global testcase_passed
        global testcase_failed

        try:
            df = pd.read_excel("result.xls", dtype='object')
            df.at[row, col] = 'Pass'
            df.to_excel("result.xls", index=False)
            testcase_passed += 1
            testcase_failed -= 1
        except Exception as e:
            print("Opps!! Something went wrong while updating excel... ", e)

    def testcase_fail(self, row=0, col=0):

        try:
            df = pd.read_excel("result.xls", dtype='object')
            df.at[row, col] = 'Fail'
            df.to_excel("result.xls", index=False)
        except Exception as e:
            print("Opps!! Something went wrong while updating excel... ", e)

    def testcase_actual_result(self, row=0, col=0, text=None):

        try:
            df = pd.read_excel("result.xls", dtype='object')
            df.at[row, col] = text
            df.to_excel("result.xls", index=False)
        except Exception as e:
            print("Opps!! Something went wrong while updating excel... ", e)