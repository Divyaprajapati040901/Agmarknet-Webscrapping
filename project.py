import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Initialize WebDriver
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)


# Load Excel data
df_commodity = pd.read_excel("Commodity_State_market.xlsx", sheet_name='Commodity')
df_state = pd.read_excel("Commodity_State_market.xlsx", sheet_name='State')
time.sleep(5) 

# URL of the website
url = 'https://www.agmarknet.gov.in/'
driver.get(url)
dropdown = WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ddlArrivalPrice"]')))
dropdown.send_keys("Both")
time.sleep(2)  # Allow time for the page to process the selection

for i in df_commodity['Commodity']:
    print({i})
    time.sleep(5)
    COMM = WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ddlCommodity"]')))
    COMM.click()
    COMM.send_keys({i})

    for j in df_state['State']:
        print({j})
        state = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ddlState"]')))
        state.send_keys({j})


        time.sleep(3)
        date_input = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/div[6]/div[1]/div/div[6]/input[1]')))
        date_input.clear()
        date_input.send_keys("01-Jan-2010")
        time.sleep(3)  # Allow time for the date to be processed    

        go_button = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/div[6]/div[1]/div/div[8]/input')))
        go_button.click()
        time.sleep(3)  # Allow time for the page to load results
    
        export_button = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/div[6]/div[6]/div[1]/div[2]/div[1]/input[2]')))
        export_button.click()
        time.sleep(3)


        # Path to the downloads directory
        source_path = r"C:\Users\DIVYA\Downloads"
        # Path to the destination directory
        destination_dir = r"E:\TADVID\commodity_crop"

        # Ensure the destination directory exists
        if not os.path.exists(destination_dir):
            os.makedirs(destination_dir)

        # Check for downloaded files
        downloaded_files = os.listdir(source_path)

        for file in downloaded_files:
            if file.endswith(".xls"):
                # Construct new filename with commodity and state names
                new_filename = f"{i}_{j}_{file}"
                # Rename the file in the source directory
                renamed_file_path = os.path.join(source_path, new_filename)
                os.rename(os.path.join(source_path, file), renamed_file_path)

                # Move the renamed file to the destination directory
                destination_path = os.path.join(destination_dir, new_filename)
                shutil.move(renamed_file_path, destination_path)
                print(f"File '{new_filename}' moved successfully to '{destination_path}'.")
                break  # Exit after processing the first matching file
