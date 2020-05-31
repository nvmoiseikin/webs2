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
import os

SIZE = 150
PAGE_START = 1
PH_FOLDER = f"ya_rings{PAGE_START}"
if not os.path.isdir(PH_FOLDER):
    os.mkdir(PH_FOLDER)
writer = pd.ExcelWriter(f'{PH_FOLDER}.xlsx', engine='xlsxwriter')
pd.DataFrame([]).to_excel(writer, sheet_name='Sheet1', index=False)

def save_img(image, index, page_number):
    print(image['src'])
    p = requests.get(f"http:{image['src'].replace(' ', '')}")
    with open(f"{PH_FOLDER}/{index + 1000*page_number}img.jpg", "wb") as f:
        f.write(p.content)
    im = Image.open(f"{PH_FOLDER}/{index + 1000*page_number}img.jpg")
    (width, height) = im.size
    print(width, height, 'РАЗМЕР')
    size = max(width, height)
    writer.sheets['Sheet1'].insert_image(101*(page_number - 1) + index + 1, 2,
            f'{PH_FOLDER}/{index + 1000*page_number}img.jpg', options={'x_scale': SIZE/size, 'y_scale': SIZE/size})

driver = webdriver.Chrome("C:/Users/Nikita/Desktop/Дима проекты/webs/chromedriver.exe")
url = "https://yandex.ru/images/search?text=rjkmwj"
driver.get(url)
content = driver.page_source
page_soup = BeautifulSoup(content, "html.parser")

urls = []

print(len(page_soup.findAll('img', attrs={'class': 'serp-item__thumb justifier__thumb'})))

for index, item_soup in enumerate(page_soup.findAll('img', attrs={'class': 'serp-item__thumb justifier__thumb'})):
    print(item_soup["src"])
    urls.append(item_soup["src"])
    save_img(item_soup, index, 1)
    time.sleep(15 + random.randint(0, 3))
    if index > 998:
        break

pd.DataFrame({"tags": urls}).to_excel(writer, sheet_name='Google_ring_tags_en', index=False)

writer.save()
writer.close()

driver.close()