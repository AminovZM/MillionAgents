from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import selenium
import time
import openpyxl
from openpyxl.styles import Font
import json

myList = []

book = openpyxl.Workbook()
sheet = book.active

sheet["A1"] = "id"; sheet.column_dimensions['A'].width = 8; sheet["A1"].font = Font(bold=True)
sheet["B1"] = "name"; sheet.column_dimensions['B'].width = 70; sheet["B1"].font = Font(bold=True)
sheet["C1"] = "url"; sheet.column_dimensions['C'].width = 65; sheet["C1"].font = Font(bold=True)
sheet["D1"] = "price"; sheet.column_dimensions['D'].width = 10; sheet["D1"].font = Font(bold=True)
sheet["E1"] = "price_promo"; sheet.column_dimensions['E'].width = 10; sheet["E1"].font = Font(bold=True)
sheet["F1"] = "brand"; sheet.column_dimensions['F'].width = 20; sheet["F1"].font = Font(bold=True)
count_excel = 2


for page in range(1, 9):
    with webdriver.Chrome() as browser:
        browser.get(f"https://online.metro-cc.ru/category/bezalkogolnye-napitki/pityevaya-voda-kulery?from=under_search&page={page}")
        time.sleep(2)
        button = browser.find_element(By.XPATH, '//*[@id="__layout"]/div/div/div[7]/div[2]/div[2]/button[1]').click()
        time.sleep(2)
        products = browser.find_element(By.ID, "products-inner").find_elements(By.CLASS_NAME, "product-card-photo__content")
        count = 1
        with webdriver.Chrome() as browser2:
            for i in products:
                print(count, i.find_element(By.TAG_NAME, "a").get_attribute("href"))
                count += 1
                browser2.get(i.find_element(By.TAG_NAME, "a").get_attribute("href"))

                try:
                    in_stock = browser2.find_element(By.CLASS_NAME, "product-page-content__prices-block").find_element(By.TAG_NAME, "p").text
                except selenium.common.exceptions.NoSuchElementException:
                    in_stock = "Yes"
                brand = browser2.find_element(By.CLASS_NAME, "product-page-content__labels-and-short-attrs").find_element(By.TAG_NAME, "ul").find_element(By.TAG_NAME, "li").find_element(By.TAG_NAME, "a").text
                article = browser2.find_element(By.CLASS_NAME, "product-page-content__article").text.split(" ")[1]
                name = browser2.find_element(By.CLASS_NAME, "product-page-content__wrapper").find_element(By.TAG_NAME, "h1").find_element(By.TAG_NAME, "span").text
                price_promo = None
                price = None
                print(in_stock)
                if in_stock != "Раскупили":
                    try:
                        price_promo = browser2.find_element(By.CLASS_NAME, "product-price-discount-above__bottom").find_element(
                            By.CLASS_NAME, "product-price__sum-rubles").text + browser2.find_element(By.CLASS_NAME,
                                                                                                     "product-price-discount-above__bottom").find_element(
                            By.CLASS_NAME, "product-price__sum-penny").text
                    except selenium.common.exceptions.NoSuchElementException:
                        price_promo = browser2.find_element(By.CLASS_NAME, "product-price-discount-above__bottom").find_element(
                            By.CLASS_NAME, "product-price__sum-rubles").text

                    try:
                        try:
                            price = browser2.find_element(By.CLASS_NAME, "product-price-discount-above__top").find_element(
                                By.CLASS_NAME, "product-price__sum-rubles").text + browser2.find_element(By.CLASS_NAME,
                                                                                                         "product-price-discount-above__top").find_element(
                                By.CLASS_NAME, "product-price__sum-penny").text
                        except selenium.common.exceptions.NoSuchElementException:
                            price = browser2.find_element(By.CLASS_NAME, "product-price-discount-above__top").find_element(
                                By.CLASS_NAME, "product-price__sum-rubles").text
                    except selenium.common.exceptions.NoSuchElementException:
                        price = " "

                if price == " ":
                    price = price_promo
                    price_promo = " "

                myList.append({"id": article, "name": name, "url": i.find_element(By.TAG_NAME, "a").get_attribute("href"), "price": price, "price_promo": price_promo, "brand": brand})

                sheet[f"A{count_excel}"] = article
                sheet[f"B{count_excel}"] = name
                sheet[f"C{count_excel}"] = i.find_element(By.TAG_NAME, "a").get_attribute("href")
                sheet[f"D{count_excel}"] = price
                sheet[f"E{count_excel}"] = price_promo
                sheet[f"F{count_excel}"] = brand
                count_excel += 1

                print(article)
                print(name)
                print(price)
                print(price_promo)
                print(brand)

json_string = json.dumps(myList)

with open("result.json", "w") as json_file:
    json_file.write(json_string)

book.save("result.xlsx")
book.close()