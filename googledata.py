# -*- coding: utf-8 -*-
"""
Created on Fri Jun 16 13:15:49 2023
@author: alpar
"""

import requests
import threading
import urllib.parse
import openpyxl
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  
chrome_options.add_argument("--disable-gpu")  

# Set up the Chrome driver service
driver_service = Service("C:/chromedriver.exe")  # Replace with the actual path to chromedriver

# Create a new instance of the Chrome driver
driver = webdriver.Chrome(service=driver_service, options=chrome_options)


def how_many_files_in_directory(path):
    count = 0
    for root_dir, cur_dir, files in os.walk(path):
        count += len(files)
    print(path)
    print('file count:', count)
    return count

# Define the function to download an image
def download_image(url, filename, path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(f"{path}{filename}", 'wb') as file:
            file.write(response.content)
            print(f"{filename} downloaded successfully.")
    else:
        print(f"{filename} download failed.")

def scrape_page(url, path, isim):
    driver.get(url)

    for i in range(3):
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        time.sleep(0.5)

    wait = WebDriverWait(driver, 15)  # Increase the timeout value here
    image_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".rg_i.Q4LuWd")))
    
    with ThreadPoolExecutor(max_workers=20) as executor:
        for idx, image_element in enumerate(image_elements):
            image_url = image_element.get_attribute("src")
            image_filename = f"google_{isim}_{idx}.jpg"
            executor.submit(download_image, image_url, image_filename, path)

# =============================================================================
# =============================================================================
# # # excel dosyası
# =============================================================================
# =============================================================================
workbook = openpyxl.load_workbook('C:/Users/hp/Desktop/Clothing_uc.xlsx')
sheet = workbook['Çeşitler']

for row in sheet.iter_rows(values_only=True):
    test = str(row)
    test_meh = test[2:-3].split(" ")

    url = f"https://www.google.com/search?q={'+'.join(test_meh)}&tbm=isch"
    path = f"C:/Users/hp/Desktop/allah/{test[2:-3]}/"

    if how_many_files_in_directory(path) < 200:
        test = test[2:-3]
        scrape_page(url, path, test)

# Quit the driver
driver.quit()