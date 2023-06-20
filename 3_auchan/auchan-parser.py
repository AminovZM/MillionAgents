from bs4 import BeautifulSoup
import requests
import time
import openpyxl
from openpyxl.styles import Font


## 20.06.2023 21:04
cookies = {
    'mindboxDeviceUUID': '5706b4ee-82d2-4a29-8538-a9c80ee56b16',
    'directCrm-session': '%7B%22deviceGuid%22%3A%225706b4ee-82d2-4a29-8538-a9c80ee56b16%22%7D',
    'isEreceiptedPopupShown_': 'true',
    'region_id': '1',
    'merchant_ID_': '1',
    'methodDelivery_': '1',
    '_GASHOP': '001_Mitishchi',
    'tmr_lvid': 'c71264c1b37c77bdb17c6c94daff3031',
    'tmr_lvidTS': '1687184522893',
    '_ym_uid': '1687184523957065910',
    '_ym_d': '1687184523',
    'rrpvid': '913975469178639',
    '_userGUID': '0:lj2y3pv8:~C9tQL5CEsaL13JOP99HPTokRQTJFNwT',
    '_ymab_param': 'wOsYcU_Ao_qYv8DtPk-DbZcESNuOUOVRP7iB-yNVrq8un-IkuZPjSU3ZKZJebmS_Gd7PAzmeNFlBhSIDIso7TUFMdKQ',
    '_ym_isad': '1',
    'rcuid': '640b35c9f3140cdce77b2d86',
    'haveChat': 'true',
    '_clck': '1o6hhq9|2|fcl|0|1265',
    '_clsk': 'wq36iw|1687184550224|1|1|q.clarity.ms/collect',
    'rrviewed': '937540%2C10390%2C606334',
    'digi_uc': 'W1sidiIsIjEwMzkwIiwxNjg3MTk1NjAxNTAwXSxbInYiLCI2MDYzMzQiLDE2ODcxODg5Njc4OTBdLFsidiIsIjkzNzU0MCIsMTY4NzE4NjgyNTgyN11d',
    'tmr_detect': '1%7C1687195601703',
    'rrlevt': '1687195601737',
    'qrator_jsr': '1687197802.767.SF6854P3Ar4gEVOC-tjqk8n2f87v4mn84rrtfjua0hkip0k66-00',
    'qrator_jsid': '1687197802.767.SF6854P3Ar4gEVOC-pivj9g7k4s8l7ugsosfbrh8obse2cld5',
    'qrator_ssid': '1687197803.952.fjkh15ze3tvOdx6Z-a78980p9l5urj7obdc9v0tf41jr3j5r6',
}

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Connection': 'keep-alive',
    # 'Cookie': 'mindboxDeviceUUID=5706b4ee-82d2-4a29-8538-a9c80ee56b16; directCrm-session=%7B%22deviceGuid%22%3A%225706b4ee-82d2-4a29-8538-a9c80ee56b16%22%7D; isEreceiptedPopupShown_=true; region_id=1; merchant_ID_=1; methodDelivery_=1; _GASHOP=001_Mitishchi; tmr_lvid=c71264c1b37c77bdb17c6c94daff3031; tmr_lvidTS=1687184522893; _ym_uid=1687184523957065910; _ym_d=1687184523; rrpvid=913975469178639; _userGUID=0:lj2y3pv8:~C9tQL5CEsaL13JOP99HPTokRQTJFNwT; _ymab_param=wOsYcU_Ao_qYv8DtPk-DbZcESNuOUOVRP7iB-yNVrq8un-IkuZPjSU3ZKZJebmS_Gd7PAzmeNFlBhSIDIso7TUFMdKQ; _ym_isad=1; rcuid=640b35c9f3140cdce77b2d86; haveChat=true; _clck=1o6hhq9|2|fcl|0|1265; _clsk=wq36iw|1687184550224|1|1|q.clarity.ms/collect; rrviewed=937540%2C10390%2C606334; digi_uc=W1sidiIsIjYwNjMzNCIsMTY4NzE4ODk2Nzg5MF0sWyJ2IiwiMTAzOTAiLDE2ODcxODkyMjA2MDhdLFsidiIsIjkzNzU0MCIsMTY4NzE4NjgyNTgyN11d; tmr_detect=1%7C1687189220825; rrlevt=1687189220899; qrator_jsr=1687191238.362.XGOHLdVb8QJwRPQs-2fa8ue1sga9gbg6rb3rpnpt5bkobuln2-00; qrator_jsid=1687191238.362.XGOHLdVb8QJwRPQs-6ob09h9uvaoq78smk1j72u3fekko4fqm',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
}


book = openpyxl.Workbook()
sheet = book.active

sheet["A1"] = "id"; sheet.column_dimensions['A'].width = 8; sheet["A1"].font = Font(bold=True)
sheet["B1"] = "name"; sheet.column_dimensions['B'].width = 70; sheet["B1"].font = Font(bold=True)
sheet["C1"] = "url"; sheet.column_dimensions['C'].width = 65; sheet["C1"].font = Font(bold=True)
sheet["D1"] = "price"; sheet.column_dimensions['D'].width = 10; sheet["D1"].font = Font(bold=True)
sheet["E1"] = "price_promo"; sheet.column_dimensions['E'].width = 10; sheet["E1"].font = Font(bold=True)
sheet["F1"] = "brand"; sheet.column_dimensions['F'].width = 20; sheet["F1"].font = Font(bold=True)
count_excel = 2

response_page = requests.get('https://www.auchan.ru/catalog/voda-soki-napitki/voda/', cookies=cookies, headers=headers)
soup_page = BeautifulSoup(response_page.text, 'lxml')

last_page = int(soup_page.find("ul", class_="css-gmuwbf").find_all("a", class_="css-jzep9t")[-2].text)
print(last_page)

for page in range(1, last_page + 1):
    print("page:", page)
    response = requests.get(f'https://www.auchan.ru/catalog/voda-soki-napitki/voda/?page={page}', cookies=cookies, headers=headers)
    soup = BeautifulSoup(response.text, 'lxml')

    links = soup.find("div", class_="css-i0ae9m css-1jbfeca-Layout").find_all("a", class_="linkToPDP active css-1kl2eos")

    count = 1
    url = ""
    for link in links:
        url = "https://www.auchan.ru" + link.get("href")
        print(count, url)
        count += 1
        response = requests.get(url, cookies=cookies, headers=headers)
        soup = BeautifulSoup(response.text, 'lxml')

        characteristics = {i.find("th").text: i.find("td").text for i in soup.find("div", class_="css-2imjyh").find("table", class_="css-p83b4h").find_all("tr")}

        article = characteristics['Артикул товара']
        name = soup.find(id="productName").text
        price = soup.find("div", class_="css-avjdfx").text.split(" ")[0]
        price_promo = ""
        if soup.find("div", class_="css-1a8h9g1") is not None:
            price_promo = soup.find("div", class_="css-1a8h9g1").text.split(" ")[0]
        brand = characteristics['Бренд']

        sheet[f"A{count_excel}"] = article
        sheet[f"B{count_excel}"] = name
        sheet[f"C{count_excel}"] = url
        sheet[f"D{count_excel}"] = price
        sheet[f"E{count_excel}"] = price_promo
        sheet[f"F{count_excel}"] = brand
        count_excel += 1
        print(article)
        print(name)
        print(price)
        print(price_promo)
        print(brand)
        # time.sleep(3)

book.save("result.xlsx")
book.close()