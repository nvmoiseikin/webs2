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

driver = webdriver.Chrome("C:/Users/Nikita/Desktop/Дима проекты/webs/chromedriver.exe")
url = "https://www.google.com/search?q=ring&tbm=isch&ved=2ahUKEwiIytXLtc3pAhXdxcQBHdQlDNsQ2-cCegQIABAA&oq=ring&gs_lcp=CgNpbWcQDDICCCkyBAgAEEMyAggAMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCABQg7UCWIO1AmDyigNoAHAAeACAAUuIAUuSAQExmAEAoAEBqgELZ3dzLXdpei1pbWc&sclient=img&ei=vOPKXoj8B92Lk74P1Muw2A0&bih=925&biw=1920&rlz=1C1GCEU_ruRU882RU885"
driver.get(url)
content = driver.page_source
page_soup = BeautifulSoup(content, "html.parser")

tags = []

print(len(page_soup.findAll('a', attrs={'class': 'KZ4CUc'})))

for index, item_soup in enumerate(page_soup.findAll('a', attrs={'class': 'F9PbJd IJRrpb'})):
    print(item_soup["aria-label"])
    tags.append(item_soup["aria-label"])

writer = pd.ExcelWriter(f'tagsEn.xlsx', engine='xlsxwriter')
pd.DataFrame({"tags": tags}).to_excel(writer, sheet_name='Google_ring_tags_en', index=False)

writer.save()
writer.close()