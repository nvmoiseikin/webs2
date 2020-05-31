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

# driver = webdriver.Chrome("C:/Users/Nikita/Desktop/Дима проекты/webs/chromedriver.exe")
# driver.get("https://www.flipkart.com/laptops/~buyback-guarantee-on-laptops-/pr?sid=6bo%2Cb5g&uniq")
# driver.get("https://www.flipkart.com/jewellery/pr?sid=mcr&q=jewellery&otracker=categorytree&page=1")
# content = driver.page_source
# soup = BeautifulSoup(content, features="html.parser")


def save_img(image, index, page_number):
    print(image[0]['src'], image[-1]['src'])
    p = requests.get(f"http:{image[-1]['src'].replace(' ', '')}")
    with open(f"фото/{index + 1000*page_number}img.jpg", "wb") as f:
        f.write(p.content)
    im = Image.open(f"фото/{index + 1000*page_number}img.jpg")
    (width, height) = im.size
    print(width, height, 'РАЗМЕР')
    size = max(width, height)
    writer.sheets['Sheet1'].insert_image(101*(page_number - 1) + index + 1, 4,
                    f'фото/{index + 1000*page_number}img.jpg', options={'x_scale': 100/size, 'y_scale': 100/size})


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

    page_soup = BeautifulSoup(response.content, "html.parser")
    for index, item_soup in enumerate(page_soup.findAll('div', attrs={'class': 'j-card-item'})):
        time.sleep(10)
        name = item_soup.find("span", {"class": "goods-name"})
        brend = item_soup.find("strong", {"class": "brand-name"})
        price = item_soup.find("ins", {"class": "lower-price"})
        image = item_soup.findAll("img", {"class": "thumbnail"})
        products.append(name.text)
        prices.append(price.text if price is not None else None)
        images.append("")
        brends.append(brend.text if brend is not None else None)
        try:
            save_img(image, index, page_number)
        except Exception as e:
            print(f"Не удалось загрузить картинку: {e}")

    df = pd.DataFrame({'Product Name': products, 'Price': prices, 'brend': brends, 'Image': images})

    print(len(df), df.iloc[0], 101*(page_number - 1) + 1)
    df.to_excel(writer, sheet_name='Sheet1', startrow=(101*(page_number - 1)), header=True, index=False)
    writer.save()
    print(page_number, " : finished")
    time.sleep(150 + random.randint(0, 20))

    if page_number < 10:
        page_number += 1
        scrapping(webpage, page_number)


writer = pd.ExcelWriter('productsW1.xlsx', engine='xlsxwriter')
pd.DataFrame([]).to_excel(writer, sheet_name='Sheet1', index=False)
url = 'https://www.wildberries.ru/catalog/yuvelirnye-ukrasheniya/zoloto?page='
scrapping(url, 1)

writer.close()
