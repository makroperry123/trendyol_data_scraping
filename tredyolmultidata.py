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


chrome_options = Options()
chrome_options.add_argument("--headless")  
chrome_options.add_argument("--disable-gpu")  

driver_service = Service("C:/chromedriver.exe") # Replace with the actual path to chromedriver


driver = webdriver.Chrome(service=driver_service, options=chrome_options)

# Define the function to download an image
def download_image(url, filename,path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(f"{path}{filename}", 'wb') as file:
            file.write(response.content)
            print(f"{filename} downloaded successfully.")
    else:
        print(f"{filename} download failed.")

# Function to scrape a page and download images in parallel
def scrape_page(url, path,isim):
    driver.get(url)
    
    for i in range(10):
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight-2250);")
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight-2350);")
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight-2450);")
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight-2550);")
        time.sleep(0.1)
    
    try:
        wait = WebDriverWait(driver, 3)
        image_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".p-card-img")))
        with ThreadPoolExecutor(max_workers=200) as executor:
            for idx, image_element in enumerate(image_elements):
                image_url = image_element.get_attribute("src")
                image_filename = f"{isim}{idx}.jpg"
                executor.submit(download_image, image_url, image_filename,path)
    except:
        print("loading error")






workbook = openpyxl.load_workbook('C:/Users/hp/Desktop/CLOTHING.xlsx')
# =============================================================================
# # excel dosyası aşağı gelmeli
# =============================================================================
sheet = workbook['ÜG4'] # 

for row in sheet.iter_rows(values_only=True):
    test = str(row)
    urle = urllib.parse.quote(test[2:-3])
    url = f"https://www.trendyol.com/sr?q={urle}&qt={urle}&st={urle}&os=1"
    path = f"C:/Users/hp/Desktop/allah/{test[2:-3]}/"
    
    if not os.path.exists(path):
        os.makedirs(path)
    test = test[2:-3]
    scrape_page(url, path,test)

# Quit the driver
driver.quit()
