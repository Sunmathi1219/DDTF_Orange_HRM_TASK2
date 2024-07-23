"""
LoginTest.py
Program : DDTF main executing file
Using DataDrivenTestingFramework,POM,Explicit wait,expected conditions,pytest kindly do the following task:
1.)Create an Excel file which will comprise of TestId,username,password,date,Timeof test,name of tester,test result for login into the portal
2)Go to the url https://opensource-demo.orangehrmlive.com/web/index.php/auth/login
3.) Login into the url using username&password provided in the Excel file.Try to user 5 username&password
4.)If the login is successful your python code will write in the Excel file whether your test passed or test failed

"""

from Locators.HomePage import OrangeHRM_Locators
from Utilities.excel_functions import SumanExcelFunctions
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
#for explicit wait only
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pytest


class Test_Login:
    excel_file = OrangeHRM_Locators().excel_file

    sheet_number = OrangeHRM_Locators().sheet_number

    # create object for the Excel Utility Class
    suman = SumanExcelFunctions(excel_file, sheet_number)

    @pytest.fixture
    #Booting function for running all pytest
    def booting_function(self):
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 10)
        yield
        self.driver.close()

    #To test the login page with 5 different set of datas and datas from excel file
    def test_login(self, booting_function):
        try:
            self.driver.get(OrangeHRM_Locators().url)

            # row count from the Excel file
            row = self.suman.row_count()

            for row in range(2, row + 1):
                username = self.suman.read_data(row, 7)
                password = self.suman.read_data(row, 8)

                user_name = self.wait.until(EC.presence_of_element_located((By.NAME, OrangeHRM_Locators().username_locator)))
                user_name.send_keys(username)

                pass_word = self.wait.until(
                    EC.presence_of_element_located((By.NAME, OrangeHRM_Locators().password_locator)))
                pass_word.send_keys(password)

                login_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, OrangeHRM_Locators().submit_button)))
                login_button.click()

                # validate the login and generate the Test-Case results & reports
                if OrangeHRM_Locators().dashboard_url == self.driver.current_url:
                    assert OrangeHRM_Locators().dashboard_url == self.driver.current_url
                    print("SUCCESS : Login with Username {a} & Password {b}".format(a=username, b=password))
                    self.suman.write_data(row, 9, OrangeHRM_Locators().pass_data)
                    self.wait.until(EC.element_to_be_clickable((By.XPATH, OrangeHRM_Locators().admin_user))).click()
                    self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT, OrangeHRM_Locators().logout_button))).click()

                elif OrangeHRM_Locators().url == self.driver.current_url:
                    assert OrangeHRM_Locators().url == self.driver.current_url
                    print("ERROR : Login unsuccessful with username {a} & Password {b}".format(a=username,
                                                                                               b=password))
                    self.suman.write_data(row, 9, OrangeHRM_Locators().fail_data)
                    self.driver.refresh()

        except NoSuchElementException as e:
            print(e)
