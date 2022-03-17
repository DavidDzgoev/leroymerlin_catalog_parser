import requests as r
import bs4 as bs
from time import time
from user_agent import generate_user_agent
import pandas as pd
import os


def download_jpg(f_name, link):
    with open(f'{f_name}', 'wb') as handle:
        response = r.get(link)

        if not response.ok:
            print(response)

        for block in response.iter_content(1024):
            if not block:
                break

            handle.write(block)


def get_excel_from_category(category_name):
    """Get .xlsx file by parsing leroymerlin.ru

    :param category_name: str
    :return: None
    """

    # get page of category
    req = r.get(
        f'https://leroymerlin.ru/catalogue/{category_name}/',
        headers={'User-Agent': generate_user_agent()}
    ).text
    soup = bs.BeautifulSoup(req, features="html.parser")

    # get number of pages
    pages = soup.find_all("a", {"class": "bex6mjh_plp o1ojzgcq_plp l7pdtbg_plp r1yi03lb_plp sj1tk7s_plp"})
    if len(pages) == 0:
        n = 1
    else:
        n = int(pages[-1].find("span", {"class": "cef202m_plp"}).text)  # number of pages

    # parsing loop
    data = []
    for i in range(1, n + 1):
        req = r.get(f'https://leroymerlin.ru/catalogue/{category_name}/?page={i}',
                    headers={'User-Agent': generate_user_agent()}).text
        soup = bs.BeautifulSoup(req, features="html.parser")

        products = soup.find_all("div", {"class": "phytpj4_plp largeCard"})  # product element

        for product in products:
            data.append(product)

    res = pd.DataFrame(columns=['Арт.', 'Фото', 'Наименование', 'Цена'])
    for product in data:
        vendor_code = product.find("span", {"class": "t3y6ha_plp sn92g85_plp p16wqyak_plp"}).text
        name = product.find("span", {"class": "t9jup0e_plp p1h8lbu4_plp"}).text
        price = product.find("p", {"class": "t3y6ha_plp xc1n09g_plp p1q9hgmc_plp"}).text
        amount = product.find("p", {"class": "t3y6ha_plp x9a98_plp pb3lgg7_plp"}).text
        img = product.find("img", {"class": "p1g8n69v_plp"}).get("src")

        res = res.append({'Арт.': vendor_code,
                          'Фото': img,
                          'Наименование': name,
                          'Цена': price + amount},
                         ignore_index=True)

    writer = pd.ExcelWriter(f'{category_name}.xlsx', engine='xlsxwriter')
    res.to_excel(writer, index=False, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']

    for index, row in res.iterrows():
        file_name = f'{row["Арт."]}.jpg'
        img_src = row["Фото"]

        download_jpg(file_name, img_src)
        worksheet.insert_image(f'B{index + 2}', file_name, {'x_offset': 2, 'y_offset': 2, 'positioning': 1})
        worksheet.write(f'B{index + 2}', " ")
        worksheet.set_row(index + 1, 100)

    worksheet.set_column('B:B', 18)

    writer.save()

    for index, row in res.iterrows():
        file_name = f'{row["Арт."]}.jpg'
        os.remove(file_name)

    return
