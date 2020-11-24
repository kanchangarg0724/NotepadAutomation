# Importing required libraries
import time
from pywinauto import *
import excel_update

# Global Variables

x1 = 0  # x coordinate of the application
y1 = 0  # y coordinate of the application

u1 = excel_update.UpdateExcel()
#update_text = "Test Case Failed due to Logical error!!"
update_text = "Testing Remote 1"

class Notepad:

    # Test Case 1: Check if the data is saved after writing to the file.
    def test_case1(self):

        x = globals()['update_text']
        try:

            app = application.Application()
            app.start("Notepad.exe")  # Open Application
            time.sleep(3)
            app.Notepad.edit.type_keys("Hello World!!", with_spaces=True)
            app.Notepad.menu_select("File->Save")
            app.Save.edit.set_text("File1.txt")
            app.SaveAs.Save.click()
            app.Notepad.close()

            # Updating result sheet
            # u1.testcase_pass(0, 'Pass/Fail')
            # x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(0, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(0, 'Actual Result', x)

    # Test Case 2: Check the functionality for the save as file.
    def test_case2(self):

        x = globals()['update_text']
        try:
            app = application.Application()
            app.start('Notepad.exe')  # Open Application
            time.sleep(3)
            app.Notepad.menu_select("File->Open")
            app.Open.edit.set_text("File1.txt")
            app.Open.type_keys('{ENTER}')
            app.Notepad.menu_select("File->Save As")
            app.SaveAs.edit.set_text("File3.txt")
            app.SaveAs.type_keys('{ENTER}')
            time.sleep(3)
            app.Notepad.close()

            # Updating result sheet
            u1.testcase_pass(1, 'Pass/Fail')
            x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(1, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(1, 'Actual Result', x)

    # Test Case 3: Check if the notepad allows same file name in the same file path.
    def test_case3(self):

        x = globals()['update_text']
        try:
            app = application.Application()
            app.start("Notepad.exe")  # Open Application
            time.sleep(3)
            app.Notepad.edit.type_keys("Hello World!!", with_spaces=True)
            app.Notepad.menu_select("File->Save")
            app.Save.edit.set_text("File1.txt")
            app.SaveAs.Save.click()
            app.ConfirmSaveAs.No.click()
            app.SaveAs.Cancel.click()
            app.Notepad.edit.type_keys('^a')
            app.Notepad.edit.type_keys('{DELETE}')
            app.Notepad.close()

            # Updating result sheet
            u1.testcase_pass(2, 'Pass/Fail')
            x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(2, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(2, 'Actual Result', x)

    # Test Case 4: Check if the window for notepad can be re-sized.
    def test_case4(self):

        global x1, y1
        x = globals()['update_text']
        try:
            app = application.Application()
            app.start("Notepad.exe")  # Open Application

            # Getting coordinates of the top left corner of a notepad window
            cords = str((app.Notepad.rectangle())).split(',')
            x1 = int(cords[0][2:])
            y1 = int(cords[1][2:])

            time.sleep(2)

            # Mouse drag event used to resize notepad window from Top Left position
            app.windows()[0].drag_mouse_input(src=(x1, y1), dst=(x1 + 20, y1 + 20))
            x1 = x1 + 20
            y1 = y1 + 20

            app.Notepad.close()
            time.sleep(3)

            # Updating result sheet
            u1.testcase_pass(3, 'Pass/Fail')
            x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(3, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(3, 'Actual Result', x)

    # Test Case 5: Check the default size of the notepad is based on last saved window size.
    def test_case5(self):

        global x1, y1
        x = globals()['update_text']
        try:
            app = application.Application()
            app.start("Notepad.exe")  # Open Application

            # Getting coordinates of the top left corner of a notepad window
            cords = str((app.Notepad.rectangle())).split(',')
            x2 = int(cords[0][2:])
            y2 = int(cords[1][2:])
            app.Notepad.close()

            if x1 == x2 and y1 == y2:
                # Updating result sheet
                u1.testcase_pass(4, 'Pass/Fail')
                x = 'Test Case Passed Successfully'

            else:
                # Updating result sheet
                u1.testcase_fail(4, 'Pass/Fail')
                x = 'Test Case Failed'

        except Exception as e:
            u1.testcase_fail(4, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(4, 'Actual Result', x)

    # Test Case 6: Check if the find option exist for the notepad.
    def test_case6(self):

        x = globals()['update_text']
        try:
            app = application.Application()
            app.start("Notepad.exe")  # Open Application
            time.sleep(3)
            app.Notepad.edit.type_keys("Hello World!!", with_spaces=True)
            app.Notepad.edit.type_keys('{HOME}^F')
            app.Find.edit.set_text("Hello")
            app.Find.FindNext.click()
            app.Find.close()
            app.Notepad.edit.type_keys('^a')
            app.Notepad.edit.type_keys('{DELETE}')
            app.Notepad.close()

            # Updating result sheet
            u1.testcase_pass(5, 'Pass/Fail')
            x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(5, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(5, 'Actual Result', x)

    # Test Case 7: Check if the notepad has find and replace functionality.
    def test_case7(self):

        x = globals()['update_text']
        try:
            app = application.Application()
            app.start("Notepad.exe")  # Open Application
            time.sleep(3)
            app.Notepad.edit.type_keys("Hello World!!", with_spaces=True)
            app.Notepad.menu_select("Edit->Replace")
            app.Replace.edit.set_text("Hello")
            app.Replace.edit.type_keys('{TAB}Bye')
            app.Replace.ReplaceAll.click()
            app.Replace.close()
            app.Notepad.edit.type_keys('^a')
            app.Notepad.edit.type_keys('{DELETE}')
            app.Notepad.close()

            # Updating result sheet
            u1.testcase_pass(6, 'Pass/Fail')
            x = 'Test Case Passed Successfully'

        except Exception as e:
            u1.testcase_fail(6, 'Pass/Fail')
            x = e
        finally:
            # Updating result sheet "Actual Result" column
            u1.testcase_actual_result(6, 'Actual Result', x)