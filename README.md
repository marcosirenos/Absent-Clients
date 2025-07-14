## Absent Clients

---

#### Objective

This is a project for automating the workflow of extracting data from flex monitor tool, process the download data, upload it to the database and refresh the Power Bi report


---

## How I’ve done it

### Web scraping

The first part of the workflow is to log in into flex monitor tool, make the desired filters and download the CSV file.

For this part I used selenium in Python, trying to make the same path as the user would with a few changes. The flex monitor tool uses frames for loading some content, for this challenge, I used a loop searching for one element through the frames, so I could interact with the website and continue the exportation. The logic is:

```python
# Setting variables
main_table_container_id = "ctl00_MainContent_ASPxSplitter1" # The element I want to find
found_in_iframe = False # Found in frame starts with false
target_iframe_element = None

iframes = driver.find_elements(By.TAG_NAME, "iframe") # Getting all frames in page

# Logic to find the element in frame
try:
    main_container_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, main_table_container_id))
    )
    logger.info(f"✅ SUCCESS: '{main_table_container_id}' found in the MAIN document!")
    found_in_iframe = True
except TimeoutException:
    logger.error(f"'{main_table_container_id}' NOT found in the main document.")
    pass
```

After all steps (more details in the file ‘extract_data_flex_monitor.py’) I simply used selenium to download the report, which will be sent to a data processing script.

---

## How to run it

For running the code, enter the cloned folder in your computer

```
cd path/to/project/folder
```

Then you can install the requirements in requirements.txt. I recommend using a VM for installing all requirements.

```
pip install -r requirements.txtt
```

Then you can run the code for the webscraping

```
python -m extract_data_flex_monitor.py
```
