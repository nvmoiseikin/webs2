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




# driver = webdriver.Chrome("C:/Users/Nikita/Desktop/Дима проекты/webs/chromedriver.exe")
# driver.get("https://www.flipkart.com/laptops/~buyback-guarantee-on-laptops-/pr?sid=6bo%2Cb5g&uniq")
# driver.get("https://www.flipkart.com/jewellery/pr?sid=mcr&q=jewellery&otracker=categorytree&page=1")
# content = driver.page_source
# soup = BeautifulSoup(content, features="html.parser")


def save_img(image, index, page_number):
    # print(image[0]['src'], image[-1]['src'])
    # p = requests.get(f"http:{image[-1]['src'].replace(' ', '')}")
    # with open(f"{PH_FOLDER}/{index + 1000*page_number}img.jpg", "wb") as f:
    #     f.write(p.content)
    try:
        im = Image.open(f"{PH_FOLDER}/{index + 1000*page_number}img.jpg")
        (width, height) = im.size
    # print(width, height, 'РАЗМЕР')
    # size = max(width, height)
        size = 336
        writer.sheets['Sheet1'].insert_image(101*(page_number%10) + index + 1, 5,
            f'{PH_FOLDER}/{index + 1000*page_number}img.jpg', options={'x_scale': SIZE/size, 'y_scale': SIZE/size})
    except Exception as e:
        print(f"Не удалось загрузить картинку: {e}")

def scrapping(webpage, page_number):
    next_page = webpage + str(page_number)
    print(next_page)
    response = requests.get(str(next_page))

    products = [] #List to store name of the product
    prices = [] #List to store price of the product
    ratings = [] #List to store rating of the product
    images = []
    descs = []
    brends = []
    urls = []
    image_urls = []

    page_soup = BeautifulSoup(response.content, "html.parser")
    for index, item_soup in enumerate(page_soup.findAll('div', attrs={'class': 'j-card-item'})):
        print(index)
        #time.sleep(6)
        name = item_soup.find("span", {"class": "goods-name"})
        brend = item_soup.find("strong", {"class": "brand-name"})
        price = item_soup.find("ins", {"class": "lower-price"})
        image = item_soup.findAll("img", {"class": "thumbnail"})
        url = item_soup.find("a", {"class": "ref_goods_n_p"})
        products.append(name.text)
        prices.append(price.text if price is not None else None)
        images.append("")
        brends.append(brend.text if brend is not None else None)
        urls.append(url['href'])
        image_urls.append(image[-1]['src'])
        try:
            save_img(image, index, page_number)
        except Exception as e:
            print(f"Не удалось загрузить картинку: {e}")

    df = pd.DataFrame({'Product Name': products, 'Price': prices,
                       'brend': brends, "url": urls, "image_url": image_urls, 'Image': images})

    print(len(df), df.iloc[0], 101*(page_number - 1) + 1)
    df.to_excel(writer, sheet_name='Sheet1', startrow=(101*(page_number%10)), header=True, index=False)
    print(page_number, " : finished")
    time.sleep(15 + random.randint(10, 20))

    if page_number < PAGE_START + 9:
        page_number += 1
        scrapping(webpage, page_number)


for page in range(3, 4):
    SIZE = 150
    PAGE_START = page*10 - 9
    PH_FOLDER = f"rings{PAGE_START}"
    if not os.path.isdir(PH_FOLDER):
        os.mkdir(PH_FOLDER)
    writer = pd.ExcelWriter(f'product_rings/products_rings{PAGE_START}.xlsx', engine='xlsxwriter')
    pd.DataFrame([]).to_excel(writer, sheet_name='Sheet1', index=False)
    url = 'https://www.wildberries.ru/catalog/yuvelirnye-ukrasheniya/zoloto/yuvelirnye-koltsa?sort=pricedown&page='

    try:
        scrapping(url, PAGE_START)
    except Exception as e:
        print(e)

    writer.save()
    writer.close()
    time.sleep(150 + random.randint(10, 20))


