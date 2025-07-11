from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import time
import os
import datetime
from auto_browser import helper_functions
import logging 
import datetime as dt

functions = helper_functions.Functions()


log_file_name = f"{dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")}_automation_log.log"
log_file_name = f"your\\log\\path\\{log_file_name}" 


logging.basicConfig(filename=log_file_name, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

user_name = "username" # Change if needed
password = "userpassword" # Change if needed


# --- Main Script Execution ---

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_argument('--ignore-certificate-errors')

driver = webdriver.Chrome(options=options)
driver.maximize_window()

logger.info("--- Running Script ---")

logger.info("1. Navigating to login page...")

driver.get("https://monitorflex.kantaribopemedia.com/Portal/Account/Login.aspx?AspxAutoDetectCookieSupport=1")

logger.info("2. Performing login...")
functions.login(driver, user_name, password)

# --- Navigate to the report page ---
logger.info("3. Accepting cookies and navigating to report search page (this page loads dynamically)...")
time.sleep(3)
functions.accept_cookies(driver)
time.sleep(3)

try:
    driver.get("https://monitorflex.kantaribopemedia.com/Portal/Portal/wfControleAbaCliente.aspx?tipopagina=relatorio&id_menu=BBBD7829-03CB-4A46-A672-6BA4F9EFBE08&relatorio=FlexivelDinamico&idModulo=77&idSessao=2&p=MonitorFlexDesconto")
    logger.info("4. Reached report search page. Now waiting for initial page load.")
except Exception as e:
    logger.error(f"ERROR: Could not navigate to report search page: {e}")
    driver.quit()
    exit()

# --- Wait for document.readyState to be 'complete' ---
print("\n5. Waiting for document.readyState to be 'complete' (up to 30 seconds)...")
try:
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    logger.info("✅ SUCCESS: document.readyState is 'complete'. Initial page scripts should be loaded.")
except TimeoutException:
    logger.warning("⚠️ WARNING: document.readyState did not become 'complete' within 30 seconds. Page might still be loading.")


# --- Logic to find the main content container by iterating through iframes ---
main_table_container_id = "ctl00_MainContent_ASPxSplitter1"
found_in_iframe = False
target_iframe_element = None

logger.info(f"\n6. Searching for iframes on the page and attempting to find '{main_table_container_id}' within each...")

time.sleep(5) # Give the page some extra time for dynamic iframes to load after initial readyState

iframes = driver.find_elements(By.TAG_NAME, "iframe")
logger.info(f"Found {len(iframes)} iframes on the page.")

# First, check the main document (default content)
logger.info("Checking default content (main document)...")
try:
    main_container_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, main_table_container_id))
    )
    logger.info(f"✅ SUCCESS: '{main_table_container_id}' found in the MAIN document!")
    found_in_iframe = True
except TimeoutException:
    logger.error(f"'{main_table_container_id}' NOT found in the main document.")
    pass

if not found_in_iframe:
    for i, iframe in enumerate(iframes):
        logger.info(f"\n    Attempting to switch to iframe {i+1}/{len(iframes)}...")
        try:
            driver.switch_to.frame(iframe)
            logger.info(f"    Switched to iframe {i+1}.")

            main_container_element_in_frame = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, main_table_container_id))
            )
            logger.info(f"✅ SUCCESS: '{main_table_container_id}' found in iframe {i+1}!")
            found_in_iframe = True
            target_iframe_element = iframe
            break

        except TimeoutException:
            logger.error(f"❌ FAILURE: '{main_table_container_id}' NOT found in iframe {i+1} within allowed time.")
        except StaleElementReferenceException:
            logger.warning(f"⚠️ WARNING: Iframe element {i+1} became stale. Skipping this iframe.")
        except Exception as e:
            logger.error(f"❌ ERROR: An unexpected error occurred while checking iframe {i+1}: {e}")
        finally:
            driver.switch_to.default_content()
            logger.info(f"    Switched back to default content from iframe {i+1}.")

# --- Perform actions only if the main container was found ---
if found_in_iframe:
    # Switch to the correct iframe for all subsequent interactions
    if target_iframe_element:
        driver.switch_to.frame(target_iframe_element)
        logger.info("\n    Switched to the target iframe for all subsequent interactions.")
    else:
        driver.switch_to.default_content()
        logger.info("\n    Interacting with elements in the main document.")


    # --- Selecting the right layout (absent-clients) ---
    # report_location = driver.find_element(By.ID, "RelatoriosDropDownList")
    # ActionChains(driver).scroll_to_element(report_location).perform()
    functions.click_element(driver, By.ID, "RelatoriosDropDownList_chosen") # Selection layouts dropdown
    locator_str = "li[data-option-array-index='3']" # The CSS location for our item 'absent-clients' layout
    functions.click_element(driver, By.CSS_SELECTOR, locator_str)
    locator_str = "//button[text()='Sim']" # Now getting the location for confirmation buttton
    functions.click_element(driver, By.XPATH, locator_str)
    # As it takes some time to load, I decided to use time.sleep() -- This can get better
    time.sleep(15)
    


    # --- 7. Select current month and year ---
    logger.info("\n7. Selecting current month and year...")
    current_year = datetime.datetime.now().year
    current_month_num = datetime.datetime.now().month
    month_names_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    current_month_name_pt = month_names_pt.get(current_month_num)

    # 7a. Select current year from "Ano" list
    logger.info("    Selecting current year...")
    year_list_css_selector = "ol.selectable.ui-selectable[argument='Ano']"
    # list_location = driver.find_element(By.CSS_SELECTOR, "By.CSS_SELECTOR, year_list_css_selector")
    # ActionChains(driver).scroll_to_element(list_location).perform()
    time.sleep(1)
    try:
        year_list_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, year_list_css_selector))
        )
        year_li_selector = f"li[data-key='{current_year}']"
        year_li_element = WebDriverWait(year_list_element, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, year_li_selector))
        )
        year_li_element.click()
        logger.info(f"    Clicked year: {current_year}.")
        time.sleep(1)
    except TimeoutException:
        logger.error(f"    ❌ FAILURE: Year {current_year} not found or not clickable in 'Ano' list. Skipping.")
    except Exception as e:
        logger.error(f"    ❌ ERROR: An error occurred while selecting year: {e}")

    # 7b. Select current month from "Mes" list
    logger.info("    Selecting current month...")
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
            logger.info(f"    Clicked month: {current_month_name_pt} (data-key: {current_month_num}).")
            time.sleep(1)
        except TimeoutException:
            logger.error(f"    ❌ FAILURE: Month '{current_month_name_pt}' ({current_month_num}) not found or not clickable in 'Mes' list. Skipping.")
        except Exception as e:
            logger.error(f"    ❌ ERROR: An error occurred while selecting month: {e}")
    else:
        logger.warning("    Could not determine current month name in Portuguese. Skipping 'Mes' selection.")


    # --- 8. Iterate through remaining target lists and Ctrl+Click all items ---
    target_list_arguments_for_all = ["UF", "Praca", "Rede", "Meio"]
    if target_list_arguments_for_all:
        logger.info(f"\n8. Iterating through remaining lists with arguments: {target_list_arguments_for_all} and clicking all items with CTRL...")
        actions = ActionChains(driver)

        for arg_name in target_list_arguments_for_all:
            logger.info(f"\n    Processing list with argument: '{arg_name}'")
            list_css_selector_for_arg = f"ol.selectable.ui-selectable[argument='{arg_name}']"

            try:
                current_list_element = WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, list_css_selector_for_arg))
                )
                logger.info(f"    List for argument '{arg_name}' located.")

                list_items = current_list_element.find_elements(By.TAG_NAME, "li")
                if not list_items:
                    logger.warning(f"    No items found in the list for argument '{arg_name}'. Skipping.")
                    continue

                logger.info(f"    Found {len(list_items)} items in the '{arg_name}' list.")

                for i in range(len(list_items)):
                    try:
                        item_selector = f"li:nth-child({i+1})"
                        item_to_click = WebDriverWait(current_list_element, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, item_selector))
                        )
                        WebDriverWait(driver, 5).until(EC.visibility_of(item_to_click))

                        actions.move_to_element(item_to_click).perform()
                        time.sleep(0.2)

                        actions.key_down(Keys.CONTROL).click(item_to_click).key_up(Keys.CONTROL).perform()
                        logger.info(f"      Clicked '{item_to_click.text}' from '{arg_name}' list with CTRL.")
                        time.sleep(0.5)
                    except TimeoutException:
                        logger.warning(f"      WARNING: Item {i+1} not found or not visible in '{arg_name}' list after re-attempt. Skipping this item.")
                        continue
                    except StaleElementReferenceException:
                        logger.warning(f"      WARNING: Item {i+1} became stale in '{arg_name}' list. Re-finding list and retrying this item...")
                        current_list_element = WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.CSS_SELECTOR, list_css_selector_for_arg))
                        )
                        list_items = current_list_element.find_elements(By.TAG_NAME, "li")
                        i -= 1
                        continue
                    except Exception as e:
                        logger.error(f"      ERROR: An error occurred while clicking item {i+1} from '{arg_name}' list: {e}")

            except TimeoutException:
                logger.error(f"    ❌ FAILURE: List with argument '{arg_name}' not found or visible within allowed time. Skipping this list.")
            except Exception as e:
                logger.error(f"    ❌ ERROR: An unexpected error occurred while processing list '{arg_name}': {e}")
    else:
        logger.warning("\nNo 'select all' lists specified or main container not found.")


    # Selecting update button to load new filters
    time.sleep(5)
    
    try:
        logger.info("  Updating table with new filters")
        functions.click_element(driver, By.ID, "AtualizarPivotGridButton")
    except Exception as e:
        logger.error(f"  Updating data error: **{e}**")

    time.sleep(15)
    
    # --- Selecting analysis tab ---
    try:
        logger.info("  Selecting analytics view")
        locator_str = "input[itemid='aba_analitico']"
        functions.click_element(driver, By.CSS_SELECTOR, locator_str)
    except Exception as e:
        logger.error(f" Fail to enter analytics view, error: **{e}**")
        
    
    time.sleep(15)
    
    # --- Finally selecting export as spreadsheet ---
    try:
        logger.info("  Exporting data")
        functions.click_element(driver, By.ID, "downloadExcel")
    except Exception as e:
        logger.error(f"  Fail export for error: **{e}**")

    # --- Final cleanup: Switch back to default content ---
    driver.switch_to.default_content()
    logger.info("\n    Switched back to default content after all interactions.")

else:
    logger.warning("\nSkipping all interactions as main container was not found in any accessible context.")

logger.info("\n--- Script finished execution. ---")
#input("Press Enter to close browser (or manually)...")
time.sleep(6)
driver.quit()