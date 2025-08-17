from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import re
import logging
import sys
from pathlib import Path
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime  # 🔹 Import for timestamp
from google.oauth2 import service_account
import pytz
import traceback
from selenium.webdriver.common.keys import Keys  
from datetime import datetime, timedelta
import calendar
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# === Setup Logging ===
# This sets up logging to the console (GitHub Actions will capture this)
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# === Setup: Linux-compatible download directory ===
download_dir = os.path.join(os.getcwd(), "download")
os.makedirs(download_dir, exist_ok=True)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")  # 🔹 Run Chrome in headless mode
chrome_options.add_argument("--disable-gpu")  # Optional: disable GPU usage
chrome_options.add_argument("--window-size=1920,1080")  # Optional: set window size for full rendering
chrome_options.add_argument("--no-sandbox")  # Optional: for Linux environments
chrome_options.add_argument("--disable-dev-shm-usage")  # Optional: prevents crashes on some systems
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

pattern = "Sales Order (sale.order)"

def is_file_downloaded():
    return any(Path(download_dir).glob(f"*{pattern}*.xlsx"))

while True:
    try:
        # === Start driver ===
        log.info("Attempting to start the browser...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 20)

        # === Step 1: Log into Odoo ===
        log.info("Navigating to login page...")
        driver.get("https://taps.odoo.com")
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys("supply.chain3@texzipperbd.com")
        driver.find_element(By.NAME, "password").send_keys("@Shanto@86")
        time.sleep(2)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(2)

        # === Step 2: Click user/company switch ===
        time.sleep(2)
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop")))
        except:
            pass

        switcher_span = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            "div.o_menu_systray div.o_switch_company_menu > button > span"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher_span)
        switcher_span.click()
        time.sleep(2)

        # === Step 3: Click 'Zipper' company ===
        log.info("Click 'Zipper' company ===")
        target_div = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//div[contains(@class, 'log_into')][span[contains(text(), 'Zipper')]]"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", target_div)
        target_div.click()
        time.sleep(2)

        # step 4
        # === Trigger global search box by sending a keystroke ===
        log.info("=== Trigger global search box by sending a keystroke ===")
        
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys("Sales")  # or use Keys.A if needed
        time.sleep(2)  # Wait for search box to appear
        
        # Step 5
        # Click on Sales option
        log.info("=== Click on Sales option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div/div[1]/a/div/span"))).click() 
        time.sleep(4)
        
        # Step 6
        
        log.info("=== Click on Filter box Down Arrow ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/button"))).click() 
        time.sleep(4)
        # Step 7
        log.info("=== Uncheck My Quotation ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/span[1]"))).click() 
        time.sleep(4)
        # Step 8
        log.info("=== Click on custome filter ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/span[11]"))).click() 
        time.sleep(4)
        
        
        # Step 9
########## 1st Conditions steps #########
        
        # Create Brance filter
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[2]/div/div[2]/button[2]"))).click() 
        time.sleep(4)
        # Add another button
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div[3]/div/div[2]/button[1]"))).click() 
        time.sleep(4)
        
        # Add another button
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div[4]/div/div[2]/button[1]"))).click() 
        time.sleep(4)
        
        # Delete the first condition
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[2]/div/div[2]/button[3]"))).click() 
        time.sleep(4)
        
        
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/div/div/div/main/div/div/div/div[2]/div/div[1]/div[1]/div/div"))).click() 
        time.sleep(4)
        
        # Step 10
        log.info("=== click on search Box and send some key like Sales Type ===")
        input_box = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/input")
        # Send the text "Sales Type"
        input_box.send_keys("Sales Team")
        time.sleep(2)
        # Step 11
        log.info("=== Click on Sales TEAM Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/ul/li/button"))).click() 
        time.sleep(3)
        # Step 12
        log.info("=== Click on  select box to find out the sales order option ===")
        # Step 13
        # EXPORT (OVERSEAS)
        input_box = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[2]/div/div[1]/div[3]/div/div/input")
        # Send the text "Sales Type"
        input_box.send_keys("EXPORT (OVERSEAS)")
        time.sleep(2)
        
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[2]/div/div[1]/div[3]/div/div/ul/li/a"))).click() 
        time.sleep(3)
        
########## Second condition Steps #########

########## Third condition Steps #########
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div/div[1]/div[1]/div/div/span"))).click() 
        time.sleep(3)
        # Step 10
        log.info("=== click on search Box and send some key like Order Date ===")
        input_box = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/input")
        # Send the text "Sales Type"
        input_box.send_keys("Order Date")
        time.sleep(2)
        # Step 11
        log.info("=== Click on order Date Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/ul/li[1]/button"))).click() 
        time.sleep(3)
        
        # Step 12
        log.info("=== Click on is to get the list of condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div/div[1]/div[2]/select"))).click() 
        time.sleep(3) 
        
        # Step 13
        log.info("=== Click on the is between condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div/div[1]/div[2]/select/option[7]"))).click() 
        time.sleep(3) 
        
        today = datetime.today()

        # === 2. Calculate the start date (first day of two months ago)
        # Adjust for year if current month is January or February
        start_month = today.month - 4
        start_year = today.year
        if start_month <= 0:
            start_month += 12
            start_year -= 1

        start_date = datetime(start_year, start_month, 1).strftime("%d/%m/%Y 00:00:21")

        # === 3. Set end date to today at 23:59:45
        end_date = today.strftime("%d/%m/%Y 23:55:45")

        # === 5. Send values to input boxes
        start_input_xpath = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div/div[1]/div[3]/div/div[1]/input"
        end_input_xpath   = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[3]/div/div[1]/div[3]/div/div[2]/input"
        time.sleep(3) 
        # === Clear and input datetime values ===
        # === Find the start input field and clear using Ctrl+A + Backspace
        start_input = driver.find_element(By.XPATH, start_input_xpath)
        start_input.send_keys(Keys.CONTROL + 'a')
        start_input.send_keys(Keys.BACKSPACE)
        start_input.send_keys(start_date)

        time.sleep(2)

        end_input = driver.find_element(By.XPATH, end_input_xpath)
        end_input.send_keys(Keys.CONTROL + 'a')
        end_input.send_keys(Keys.BACKSPACE)
        end_input.send_keys(end_date)
        time.sleep(2)
       
       
       ### Step 3 TO ADD 3RD PLUS BUTTON

        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[4]/div/div[1]/div[1]/div/div/span"))).click() 
        time.sleep(3)
        # Step 10
        log.info("=== click on search Box and send some key like Order Date ===")
        input_box = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/input")
        # Send the text "Sales Type"
        input_box.send_keys("Sales Type")
        time.sleep(2)
        # Step 11
        log.info("=== Click on sales Type Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/ul/li/button"))).click() 
        time.sleep(3)
        
        
        # Step 13
        log.info("=== Click on the is between condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[4]/div/div[1]/div[3]/select"))).click() 
        time.sleep(3) 
       
        log.info("=== Click on the is between condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[4]/div/div[1]/div[3]/select/option[5]"))).click() 
        time.sleep(3) 
        
        
        ### Step 12
            
        log.info("=== Click on custome filter first search box salesperson to change it value ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[5]/div/div[1]/div[1]/div/div/span"))).click() 
        time.sleep(3)
        
        log.info("=== click on search Box and send some key like Order Date ===")
        input_box = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/input")
        # Send the text "Sales Type"
        input_box.send_keys("Status")
        time.sleep(2)
        # Step 11
        log.info("=== Click on sales Type Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/ul/li[1]/button"))).click() 
        time.sleep(3)
        
        
        # Step 13
        log.info("=== Click on the is between condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[5]/div/div[1]/div[3]/select"))).click() 
        time.sleep(3) 
       
        log.info("=== Click on the is between condition Option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[5]/div/div[1]/div[3]/select/option[3]"))).click() 
        time.sleep(3) 
        
        
       
       ## Click on any to all
        log.info("=== Click on Any to All option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[1]/div/div/div/button"))).click() 
        time.sleep(2) 
        log.info("=== Click on  All option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div[1]/div/div/div/div/span[1]"))).click() 
        time.sleep(2) 
        
        # Click on confirm or loading data 
        log.info("=== Click on  Confirm option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/button[1]"))).click() 
        time.sleep(10) 
        
       
 ######## condition step is completed ########
 
 ####### Now file downloading steps ########## 
 
        # Click on checkbox to select all the data 
        log.info("=== Click on  checkbox option ===")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[3]/div[2]/table/thead/tr/th[1]/div"))).click() 
        time.sleep(2) 
            # Step 2: Check if "Select All" text is present anywhere on the page
        if "Select all" in driver.page_source:
            log.info("=== 'Select all' text found. Proceeding with normal flow ===")

            # Click on "Select All"
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[1]/span/a[1]"))).click()
            time.sleep(2)
            # Continue with action → export → download flow
            log.info("=== Click on Action option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/button"))).click()
            time.sleep(2)

            log.info("=== Click on export option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/div/span[1]"))).click()
            time.sleep(2)

            log.info("=== Click on download template option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select"))).click()
            time.sleep(2)

            log.info("=== Select custom template 0_ABCD_arifuls ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select/option[2]"))).click()
            time.sleep(2)

            log.info("=== Final export button click ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/button[1]"))).click()
            time.sleep(30)
        else:
            # Continue with action → export → download flow
            log.info("=== Click on Action option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/button"))).click()
            time.sleep(2)

            log.info("=== Click on export option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/div/span[1]"))).click()
            time.sleep(2)

            log.info("=== Click on download template option ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select"))).click()
            time.sleep(2)

            log.info("=== Select custom template 0_ABCD_arifuls ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select/option[2]"))).click()
            time.sleep(2)

            log.info("=== Final export button click ===")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/button[1]"))).click()
            time.sleep(30)
       
        # === Step 9: Confirm file downloaded ===
        
        # === Step 9: Confirm file downloaded ===
        if is_file_downloaded():
            log.info("✅ File download complete!")
            files = list(Path(download_dir).glob(f"*{pattern}*.xlsx"))
            if len(files) > 1:
                files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                for file in files[1:]:
                    file.unlink()
            driver.quit()
            break  # Exit the loop after file download is complete
        else:
            log.warning("⚠️ File not downloaded. Retrying...")
            

    except Exception as e:
        driver.save_screenshot("error_screenshot.png")
        log.error(f"❌ Error Roccurred: {traceback.format_exc()}\nRetrying in 10 seconds...\n")
        try:
            driver.quit()
        except:
            pass
        time.sleep(10)
        

# === Step 11: Load latest file and paste to Google Sheet ===
try:
    log.info("Checking for downloaded files...")
    files = list(Path(download_dir).glob(f"*{pattern}*.xlsx"))
    if not files:
        raise Exception("No matching file found.")

    # Sort and get the latest file
    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    latest_file = files[0]
    print(f"Latest file found: {latest_file.name}")

    # Load into DataFrame
    df_production_pcs = pd.read_excel(latest_file,sheet_name=0)
    print("File loaded into DataFrame.")
    
    # Setup Google Sheets API
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=scope)
    log.info("✅ Successfully loaded credentials.")

    # Use gspread to authorize and access Google Sheets
    client = gspread.authorize(creds)

    # Open the sheet and paste the data
    sheet = client.open_by_key("1l2xcuZVCgj3yVVKerFE9iCIK1SvyHUMWpZQ5af5wbLM")
    worksheet = sheet.worksheet("Export Overseas OA Data")
    
    if df_production_pcs.empty:
        print("Skip: DataFrame is empty, not pasting to sheet.")
    else:
    # Clear old content (optional)
        worksheet.batch_clear(['A:AC'])
        # Paste new data
        set_with_dataframe(worksheet, df_production_pcs)
        print("Data pasted to Google Sheet (Sheet4).")
        # === ✅ Add timestamp to Y2 ===
        local_tz = pytz.timezone('Asia/Dhaka')
        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update("AC2", [[f"{local_time}"]])
        print(f"Timestamp written to AC2: {local_time}")
    

except Exception as e:
    print(f"Error while pasting to Google Sheets: {e}")
