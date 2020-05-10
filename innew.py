import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
import openpyxl
from bs4 import BeautifulSoup
import time

isbn = []
seller = []
bol_price = []

c = open("learn1.txt", "r")

options = Options()
options.add_argument("window-size=1400,600")
ua = UserAgent()
a = ua.random
user_agent = ua.random

options.add_argument(f'user-agent={user_agent}')
driver = webdriver.Chrome("D:\game\ok\chromedriver.exe", options=options)

note = 0
bolloop = 0
isbnnumber = 0
selup = 0

bol = 'https://www.bol.com/nl/s/?searchtext='

for no in c:
    isbnnumber = no
    if no == c:
        break
c.close()

c = open("learn1.txt", "r")

for esb in c:
    note = note + 1
    try:
        isbn.append(esb)
    except:
        print("Done")
    if note == isbnnumber:
        break
c.close()

c = open("learn1.txt", "r")

for isbnn in c:
    time.sleep(1)
    bolloop = bolloop + 1
    url = bol + isbnn
    driver.implicitly_wait(30)
    driver.get(url)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    zakum = soup.find_all("div", {"class": "price-block__highlight"})
    if zakum is False or zakum is None or zakum == "" or zakum == '' or zakum == []:
        bol_price.append('Nothing')
    else:
        fa = soup.find_all("meta", {"itemprop": "price"})
        for abc in fa:
            bol_price.append(abc['content'])
    if bolloop == isbnnumber:
        break
c.close()

c = open("learn1.txt", "r")

for nice in c:
    time.sleep(1)
    selup = selup + 1
    trl = bol + nice
    driver.implicitly_wait(30)
    driver.get(trl)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    sell = soup.find_all("div", {"class": "product-seller medium--is-visible"})
    if sell is False or sell is None or sell == "" or sell == '' or sell == []:
        seller.append('Nothing')
    else:
        ta = soup.find_all("div", {"class": "product-seller medium--is-visible"})
        for abz in ta:
            seller.append(abz.get_text())
    if selup == isbnnumber:
        break
c.close()

z = {"ISBN": isbn, "SELLER": seller, "bolprice": bol_price}
dfs = pd.DataFrame.from_dict(z, orient='index')
z = dfs.transpose()
af = z.to_excel('output.xlsx')
