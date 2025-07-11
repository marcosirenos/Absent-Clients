# functions.py

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

class Functions:

    @staticmethod
    def click_element(driver, by_type, locator, timeout=10):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by_type, locator))
            )
            element.click()
            print(f"Clicked: {locator}")
        except Exception as e:
            print(f"ERROR: Could not click {locator}. Exception: {e}")
            raise

    @staticmethod
    def fill_input(driver, by_type, locator, text, timeout=10):
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by_type, locator))
            )
            element.clear()
            element.send_keys(text)
            print(f"Filled: {locator}")
        except Exception as e:
            print(f"ERROR: Could not fill {locator}. Exception: {e}")
            raise

    @staticmethod
    def login(driver, username, password):
        Functions.fill_input(driver, By.ID, "ctl00_UserName", username)
        Functions.fill_input(driver, By.ID, "ctl00_Password", password)
        Functions.click_element(driver, By.ID, "ctl00_btnLogin")

    @staticmethod
    def find_element_with_wait(driver, by_type, locator, timeout=15):
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by_type, locator))
        )

    @staticmethod
    def find_clickable_element_with_wait(driver, by_type, locator, timeout=15):
        return WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by_type, locator))
        )

    @staticmethod
    def double_click_element(driver, element_id, timeout=10):
        try:
            element = driver.find_element(By.ID, element_id)
            ActionChains(driver).double_click(element).perform()
            print(f"Double-clicked folder to expand: {element}")
            time.sleep(3)
        except Exception as e:
            print(f"No element found, double click failed: {e}")

    @staticmethod
    def accept_cookies(driver, timeout=5):
        driver.implicitly_wait(timeout)
        try:
            Functions.click_element(driver, By.ID, "CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll")
            print("   Successfully accepted cookies.")
        except Exception as e:
            print(f"  Failed to accept cookies, error: **{e}**")