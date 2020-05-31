from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from StyleFrame import StyleFrame, Styler
import requests
import io
import time
import xlsxwriter
import random
from PIL import Image
SIZE = 175
PAGE_START = 21
PH_FOLDER = f"rings{PAGE_START}"


def save_img(index, page_number):
    try:
        im = Image.open(f"{PH_FOLDER}/{index + 1000*page_number}img.jpg")
        (width, height) = im.size
        print(index, page_number, 'indexes')
        size = max(width, height)
        writer.sheets['Sheet1'].insert_image((((page_number-1)%10)*100 + index)//10, (index%10) * 2,
                f'{PH_FOLDER}/{index + 1000*page_number}img.jpg', options={'x_scale': SIZE/size, 'y_scale': SIZE/size})
    except Exception as e:
        print(f"Не удалось загрузить картинку: {e}")


writer = pd.ExcelWriter(f'photos_excel/photos{PAGE_START}.xlsx', engine='xlsxwriter')
pd.DataFrame([]).to_excel(writer, sheet_name='Sheet1', index=False)
for page_number in range(PAGE_START, PAGE_START+10):
    for index in range(100):
        save_img(index, page_number)

writer.save()
writer.close()
