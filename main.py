import os
import sys
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
import requests
import time

# 2017219944
root_path = "https://patentscope2.wipo.int/search/en/search.jsf"


def getRequestDatas():
    file = "Patents_List.xlsx"

    #load the work book
    wb_obj = load_workbook(filename = file)
    wsheet = wb_obj['20220326']
    dataDict = {}
    row_count = wsheet.max_row
    col_count = wsheet.max_column
    print("row_count",row_count)
    print("col_count",col_count)

    for row in range(2, row_count+1):
        key = wsheet.cell(row=row, column=3).value
        value = wsheet.cell(row=row, column=4).value
        dataDict[key] = value
    
    return dataDict


def scrap_one_page(dataDict):
    # foreach dataDict
    for key, value in dataDict.items():
        searchString = key
        print("searchString",searchString)
        get_input_search_click_button(searchString)


def get_input_search_click_button(searchString):

    options = webdriver.ChromeOptions()
    options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    driver = webdriver.Chrome(chrome_options=options,executable_path=r'D:\\develop_experiences\\py_scraper_patentscope\\chromedriver.exe')
    time.sleep(10)
    body_data = driver.get(root_path)
    body_content = BeautifulSoup(body_data, 'html.parser')
 

def start():

    dataDict=getRequestDatas()
    print("dataDict",dataDict)
    scrap_one_page(dataDict)


start()
