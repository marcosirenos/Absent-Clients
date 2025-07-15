import os
import time
import datetime
import logging
import shutil
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from functions import helper_functions

class WebScraping:
    def __init__(self):
        self.functions = helper_functions.Functions()
        self.logger = self.set_log()
        self.download_dir = self.check_directories()
        self.driver = None
        self.target_iframe_element = None

    @staticmethod
    def set_log():
        current_working_directory = Path.cwd()
        log_directory = current_working_directory / "log"
        log_directory.mkdir(parents=True, exist_ok=True)
        log_file_name = f"{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}_scraping.log"
        log_file_path = log_directory / log_file_name
        logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        logger = logging.getLogger()
        return logger

    @staticmethod
    def check_directories():
        script_dir = os.path.dirname(os.path.abspath(__file__))
        download_dir = os.path.join(script_dir, os.pardir, "data", "raw")
        download_dir = os.path.abspath(download_dir)
        if os.path.exists(download_dir):
            print(f"Clearing contents of folder: {download_dir}")
            for filename in os.listdir(download_dir):
                file_path = os.path.join(download_dir, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                    print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Failed to delete {file_path}. Reason: {e}")
        else:
            os.makedirs(download_dir)
            print(f"Created destination folder: {download_dir}")
        return download_dir

    def initialize_driver(self):
        self.logger.info("Initializing Chrome driver")
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("detach", True)
        options.add_argument('--ignore-certificate-errors')
        try:
            self.driver = webdriver.Chrome(options=options)
        except Exception as e:
            self.logger.error(f"Error when trying to initialize Chrome driver -- {e}")
            raise
        self.logger.info("Driver set up")

    def check_frames(self, main_element_id):
        self.logger.info(f"Searching iframes on the page attempting to find {main_element_id}")
        time.sleep(2)
        driver = self.driver
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        self.logger.info(f"Found {len(iframes)} iframes on  the page")
        self.logger.info("Checking default content")
        found_in_frame = False
        try:
            main_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, main_element_id))
            )
            self.logger.info(f"SUCCESS: {main_element_id} found in MAIN document")
            found_in_frame = True
            self.target_iframe_element = None
        except TimeoutException:
            self.logger.info(f"{main_element_id} NOT found in main document")
            pass
        if not found_in_frame:
            for i, iframe in enumerate(iframes):
                self.logger.info(f"Attempting to switch to iframe {i+1}/{len(iframes)}")
                try:
                    driver.switch_to.frame(iframe)
                    self.logger.info(f"Switched to iframe {i+1}")
                    main_container_element_in_frame = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.ID, main_element_id))
                    )
                    self.logger.info(f"SUCCESS: '{main_element_id}' found in frame")
                    found_in_frame = True
                    self.target_iframe_element = iframe
                    break
                except TimeoutException:
                    self.logger.error(f"FAILURE {main_element_id} not found in frame {i+1}")
                except StaleElementReferenceException:
                    self.logger.warning(f"WARNING iframe {i+1} became stale")
                except Exception as e:
                    self.logger.error(f"ERROR an unexpected error occurred: {e}")
                finally:
                    driver.switch_to.default_content()
                    self.logger.info(f"Switched back to default content from frame {i+1}")
        return found_in_frame

    def web_scraping(self, user_name, password):
        self.initialize_driver()
        driver = self.driver
        main_element_id = "ctl00_MainContent_ASPxSplitter1"
        self.logger.info("Navigating to login page")
        driver.get("https://monitorflex.kantaribopemedia.com/Portal/Account/Login.aspx?AspxAutoDetectCookieSupport=1")
        self.logger.info("Performing login")
        self.functions.login(driver, user_name, password)
        self.logger.info("Accepting cookies and navigating to report search page")
        time.sleep(3)
        self.functions.accept_cookies(driver)
        time.sleep(3)
        try:
            driver.get("https://monitorflex.kantaribopemedia.com/Portal/Portal/wfControleAbaCliente.aspx?tipopagina=relatorio&id_menu=BBBD7829-03CB-4A46-A672-6BA4F9EFBE08&relatorio=FlexivelDinamico&idModulo=77&idSessao=2&p=MonitorFlexDesconto")
            self.logger.info("Reached report search page")
        except Exception as e:
            self.logger.error(f"ERROR: Could not navigate to report search page: {e}")
            driver.quit()
            return
        self.logger.info("Waiting for page to be ready")
        try:
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            self.logger.info("SUCCESS: document.readyState is 'complete'. Initial page scripts should be loaded.")
        except TimeoutException:
            self.logger.warning("WARNING: document.readyState did not become 'complete' within 30 seconds. Page might still be loading.")
        self.logger.info("Checking iframes")
        time.sleep(3)
        found_in_frame = self.check_frames(main_element_id)
        if found_in_frame:
            if self.target_iframe_element:
                driver.switch_to.frame(self.target_iframe_element)
                self.logger.info("Switched to target iframe")
            else:
                driver.switch_to.default_content()
                self.logger.info("Interacting with elements in the main document")
            self.functions.click_element(driver, By.ID, "RelatoriosDropDownList_chosen")
            locator_str = "li[data-option-array-index='3']"
            self.functions.click_element(driver, By.CSS_SELECTOR, locator_str)
            locator_str = "//button[text()='Sim']"
            self.functions.click_element(driver, By.XPATH, locator_str)
            time.sleep(15)
            self.logger.info("Selecting current month and year")
            current_year = datetime.datetime.now().year
            current_month_num = datetime.datetime.now().month
            month_names_pt = {
                1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
                7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
            }
            current_month_name_pt = month_names_pt.get(current_month_num)
            self.logger.info("Selecting current year")
            year_list_css_selector = "ol.selectable.ui-selectable[argument='Ano']"
            try:
                year_list_element = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, year_list_css_selector))
                )
                year_li_selector = f"li[data-key='{current_year}']"
                year_li_element = WebDriverWait(year_list_element, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, year_li_selector))
                )
                year_li_element.click()
                self.logger.info(f"Clicked year: {current_year}")
                time.sleep(1)
            except TimeoutException:
                self.logger.error(f"FAILURE: Year {current_year} not found or not clickable in 'Ano' list. Skipping.")
            except Exception as e:
                self.logger.error(f"ERROR: An error occurred while selecting year: {e}")
            self.logger.info("Selecting current month")
            month_list_css_selector = "ol.selectable.ui-selectable[argument='Mes']"
            if current_month_name_pt:
                try:
                    month_list_element = WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, month_list_css_selector))
                    )
                    month_li_selector = f"li[data-key='{current_month_num}']"
                    month_li_element = WebDriverWait(month_list_element, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, month_li_selector))
                    )
                    month_li_element.click()
                    self.logger.info(f"Clicked month: {current_month_name_pt} (data-key: {current_month_num})")
                    time.sleep(1)
                except TimeoutException:
                    self.logger.error(f"FAILURE: Month '{current_month_name_pt}' ({current_month_num}) not found or not clickable in 'Mes' list. Skipping.")
                except Exception as e:
                    self.logger.error(f"ERROR: An error occurred while selecting month: {e}")
            else:
                self.logger.warning("Could not determine current month name in Portuguese. Skipping 'Mes' selection.")
            target_list_arguments_for_all = ["UF", "Praca", "Rede", "Meio"]
            if target_list_arguments_for_all:
                self.logger.info(f"Iterating through remaining lists with arguments: {target_list_arguments_for_all} and clicking all items with CTRL")
                actions = ActionChains(driver)
                for arg_name in target_list_arguments_for_all:
                    self.logger.info(f"Processing list with argument: '{arg_name}'")
                    list_css_selector_for_arg = f"ol.selectable.ui-selectable[argument='{arg_name}']"
                    try:
                        current_list_element = WebDriverWait(driver, 20).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, list_css_selector_for_arg))
                        )
                        self.logger.info(f"List for argument '{arg_name}' located")
                        list_items = current_list_element.find_elements(By.TAG_NAME, "li")
                        if not list_items:
                            self.logger.warning(f"No items found in the list for argument '{arg_name}'. Skipping")
                            continue
                        self.logger.info(f"Found {len(list_items)} items in the '{arg_name}' list")
                        for i in range(len(list_items)):
                            try:
                                item_selector = f"li:nth-child({i+1})"
                                item_to_click = WebDriverWait(current_list_element, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, item_selector))
                                )
                                WebDriverWait(driver, 5).until(EC.visibility_of(item_to_click))
                                actions.move_to_element(item_to_click).perform()
                                actions.key_down(Keys.CONTROL).click(item_to_click).key_up(Keys.CONTROL).perform()
                                self.logger.info(f"Clicked '{item_to_click.text}' from '{arg_name}' list with CTRL.")
                            except TimeoutException:
                                self.logger.warning(f"Item {i+1} not found or not visible in '{arg_name}' list after re-attempt. Skipping this item")
                                continue
                            except StaleElementReferenceException:
                                self.logger.warning(f"Item {i+1} became stale in '{arg_name}' list. Re-finding list and retrying this item...")
                                current_list_element = WebDriverWait(driver, 10).until(
                                    EC.visibility_of_element_located((By.CSS_SELECTOR, list_css_selector_for_arg))
                                )
                                list_items = current_list_element.find_elements(By.TAG_NAME, "li")
                                continue
                            except Exception as e:
                                self.logger.error(f"ERROR: An error occurred while clicking item {i+1} from '{arg_name}' list: {e}")
                    except TimeoutException:
                        self.logger.error(f"FAILURE: List with argument '{arg_name}' not found or visible within allowed time. Skipping this list.")
                    except Exception as e:
                        self.logger.error(f"ERROR: An unexpected error occurred while processing list '{arg_name}': {e}")
            else:
                self.logger.warning("No 'select all' lists specified or main container not found.")
            try:
                self.logger.info("  Updating table with new filters")
                self.functions.click_element(driver, By.ID, "AtualizarPivotGridButton")
            except Exception as e:
                self.logger.error(f"  Updating data error: **{e}**")
            time.sleep(15)
            try:
                self.logger.info("  Selecting analytics view")
                locator_str = "input[itemid='aba_analitico']"
                self.functions.click_element(driver, By.CSS_SELECTOR, locator_str)
            except Exception as e:
                self.logger.error(f" Fail to enter analytics view, error: **{e}**")
            time.sleep(15)
            try:
                self.logger.info("  Exporting data")
                self.functions.click_element(driver, By.ID, "downloadExcel")
            except Exception as e:
                self.logger.error(f"  Fail export for error: **{e}**")
            driver.switch_to.default_content()
            self.logger.info("\n    Switched back to default content after all interactions.")
        else:
            self.logger.warning("\nSkipping all interactions as main container was not found in any accessible context.")
        self.logger.info("Script finished execution")
        time.sleep(6)
        driver.quit()