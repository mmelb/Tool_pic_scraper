from bs4 import BeautifulSoup
from selenium import webdriver
import re
import time
import requests
from openpyxl import load_workbook


def get_tools():
    nos = {}
    i = 2
    while True:
        if ws[f"a{i}"].value:
            nos.update({ws[f"a{i}"].value: ws[f"r{i}"].value})
        if not ws[f"a{i}"].value:
            break
        i += 1
    return nos


def scrape_and_download(art, link, line):
    site = link
    driver = webdriver.Chrome()
    driver.get(site)
    time.sleep(3)

    pics = BeautifulSoup(driver.page_source, 'html.parser')
    img_tags = pics.find_all('img')

    urls = [img['src'] for img in img_tags]

    for url in urls:
        filename = \
            re.search(rf'{line}.jpg$', url)
        if filename:
            print(f"Found picture for: {art}, downloading...")
            with open(f'guhring/{art}.gif', 'wb') as handle:
                response = requests.get(url, stream=True)
                if not response.ok:
                    print(response)
                for block in response.iter_content(1024):
                    if not block:
                        break

                    handle.write(block)
            break


def gen_link(line, dim):
    split_dim = str(dim).split(".")
    dim = str(split_dim[0].zfill(3)) + str(split_dim[1]) + ((4 - len(split_dim[1])) * "0")
    return f"https://webshop.guehring.no/0000090{str(line).zfill(4)}{dim}"


# wb = load_workbook('Coroturn_CXS.xlsx', data_only=True)
# ws = wb["Ark1"]
art = 124153
line = 391
dim = 2.184
print(gen_link(line, dim))

# art_nos = get_tools()

# for key in art_nos:
scrape_and_download(art, gen_link(line, dim), line)
