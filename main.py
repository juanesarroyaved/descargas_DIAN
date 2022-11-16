# -*- coding: utf-8 -*-

import os
os.chdir(r"C:\Z_Proyectos\descargas_DIAN")

import config
import pdf_tables as pt
import pandas as pd
from glob import glob
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

def initialize_driver():
    global driver
    
    options = Options()
    options.page_load_strategy = 'eager'
    options.add_argument('start-maximized')
    options.add_argument('enable-automation')
    options.add_argument('--disable-browser-side-navigation')
    options.add_argument('--disable-gpu')
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
def quit_driver():
    driver.quit()
    
def dian_login(params: dict = config.params):
    driver.get(params['url'])
     
def search_docs(params: dict = config.params):
    driver.maximize_window()
    driver.get(config.url_received)
    field_dates = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="dashboard-report-range"]')))
    field_dates.clear()
    field_dates.send_keys(params['dates_str'], Keys.ENTER)

    search_button = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="document-list"]/div[2]/button')))
    search_button.click()

    select_menu = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="tableDocuments_length"]/label/select')))
    select_menu = Select(select_menu)
    select_menu.select_by_value('100')

def download_docs():
    
    num_pages = WebDriverWait(driver, 15).until(EC.presence_of_all_elements_located((By.XPATH, r'//*[@id="tableDocuments_paginate"]/span/a')))
    
    for page in num_pages:
        page.click()
        download_buttons = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, r'//*[@class="btn btn-xs add-tooltip download-document"]')))
        for button in download_buttons:
            cufe = button.get_attribute('data-id')
            print(cufe)
            retrieve_http(cufe)
            
def retrieve_http(cufe):
    driver.get(config.download_doc_url + cufe)
    time.sleep(2)
    
def doc_list_from_file(doc_list_path: str = config.doc_list_path):
    df_docs = pd.read_excel(doc_list_path)
    cufes = df_docs['CUFE/CUDE'].unique()
    
    for cufe in cufes:
        retrieve_http(cufe)
    

def dian_pipeline(params: dict = config.params):
    zips_glob = config.download_path + '*.zip'
    not_zips = glob(zips_glob)
    initialize_driver()
    dian_login(params)
    
    if params['From_xlsx'] == 'SI' and os.path.isfile(params['xlsx_path']):
        doc_list_from_file()
    else:
        search_docs(params)
        download_docs()
    
    quit_driver()
    
    zips = [i for i in glob(zips_glob) if i not in not_zips]
    
    return zips
    
def lector_pdf_pipeline(zips: list, export: bool = True):
    for z in zips:
        pt.move_file(z, config.zips_paths)
    df_total = pt.main(export=True)
    
    return df_total

if __name__ == '__main__':
    zips = dian_pipeline()
    df_total = lector_pdf_pipeline(zips)
    
