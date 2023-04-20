from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
from bs4 import BeautifulSoup
import time
from os import path, getcwd, listdir
import glob
import os, os.path
from PIL import Image
from io import BytesIO
import traceback
#import requests
import urllib.request as request
from webdriver_manager.chrome import ChromeDriverManager
pd.set_option('display.max_colwidth', 200)



PATH_APP = getcwd()
PATH_OUT_DIR = path.join(PATH_APP, 'out')
path_img = r'C:\Users\User\Desktop\Regression\Carros\captchas\sunarp'
qs = 300
ejemplo = {'Oficina':['LIMA'], 'Año':['2022'], 'Titulo':['2510798']}
ejemplo = pd.DataFrame.from_dict(ejemplo)

def extract_captcha(driver, q):
    
    ref_img = driver.find_element_by_id('image')
    driver.execute_script("window.scrollTo(0,0);")

    location = ref_img.location
    size = ref_img.size
    png = driver.get_screenshot_as_png()
    try:
        with Image.open(BytesIO(png)) as img:
            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']
            img = img.crop((left, top, right, bottom))
            name = 'sunarp_{}.png'.format(q)
            path_captcha = path.join(path_img, name)
            img.save(path_captcha)
            print('Image {} saved'.format(name))
        return path_captcha
    except Exception as e:
        print(traceback.format_exc())
        raise Exception(f'Error saving captcha image {e}')

def change_window(driver, index):
    windows = driver.window_handles
    if (index < len(windows)):
        driver.switch_to.window(windows[index])

def log(office, year, title):

    options = webdriver.ChromeOptions()
    options.add_experimental_option('prefs', {
    "download.default_directory": r"C:\Users\User\Desktop\Regression\Carros\pdf", #Change default directory for downloads
    "download.prompt_for_download": False, #To auto download the file
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
    })

    URL = 'https://siguelo.sunarp.gob.pe/siguelo/'
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.get(URL)
    driver.maximize_window()
    off = driver.find_element_by_xpath('/html/body/app-root/app-ingreso/body/div/div/div/div/div/form/div[2]/div/select')
    office_options = off.find_elements_by_tag_name('option')
    dict_office_options = { i.text : i for i in office_options }
    dict_office_options[office].click()

    years = driver.find_element_by_xpath('/html/body/app-root/app-ingreso/body/div/div/div/div/div/form/div[3]/div/div/div/select')
    year_options = years.find_elements_by_tag_name('option')
    dict_year_options = { i.text : i for i in year_options }
    dict_year_options[year].click()

    title_input = driver.find_element_by_xpath('/html/body/app-root/app-ingreso/body/div/div/div/div/div/form/div[4]/div/div/input')
    title_input.send_keys(title)

    ## solve captcha function
    time.sleep(6)

    search_bttn = driver.find_element_by_xpath('/html/body/app-root/app-ingreso/body/div/div/div/div/div/form/div[6]/div/div/button')
    search_bttn.click()
    time.sleep(6)

    access = driver.find_element_by_xpath('/html/body/app-root/app-titulo/body/div[11]/div[3]/table/tr[2]/td/a')
    access.click()
    time.sleep(4)

    open_file = driver.find_element_by_xpath('/html/body/div[2]/div[2]/div/mat-dialog-container/app-partidas/mat-card-content/div/div/table/tbody/tr/td[3]/button')
    open_file.click()
    time.sleep(6)

    filename = '{}.pdf'.format(title)
    list_of_files = glob.glob(r'C:\Users\User\Desktop\Regression\Carros\pdf\*') # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    full_path = os.path.join(PATH_OUT_DIR, filename)
    os.rename(latest_file, full_path)

    driver.quit()

asd = ejemplo.apply(lambda x: log(x['Oficina'], x['Año'], x['Titulo']), axis=1)


