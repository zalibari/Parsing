import requests
from bs4 import BeautifulSoup
from time import sleep
import xlsxwriter

headers = {"User-Agent":
               "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.5) Gecko/20091102 Firefox/3.5.5 (.NET CLR 3.5.30729)"}


def get_url():
    for count in range(1, 6):
        url = f"https://www.ruscable.ru/company/company_rtl/cable_comp/?page={count}"
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, "lxml")
        data = soup.find_all('tr', class_="table_rtl_body")
        for i in data:
            card_url = i.find('a').get('href')
            country = i.find('span', class_="comp_location").text.replace('\n', '').strip()
            reputation = i.find('td', class_='comp_trl').text.replace('\n', '').strip()
            yield card_url, country, reputation


def array():
    for card_url, country, reputation in get_url():
        response = requests.get(card_url, headers=headers)
        sleep(2)
        soup = BeautifulSoup(response.text, "lxml")
        data = soup.find('h1').text.replace('\n', '').strip()
        try:
            address = soup.find('span', itemprop="streetAddress").text.replace('\n', '').strip()
            telephone = soup.find('span', itemprop="telephone").text.replace('\n', '').strip()
            mail = soup.find('span', itemprop="email").text.replace('\n', '').strip()
            site = soup.find('span', itemprop="url").text.replace('\n', '').strip()
        except Exception:
            try:
                company_id = soup.find('div', class_="contacts").find('a').attrs['comp_id']
                popup_url = f'https://www.ruscable.ru/company/contacts.php?comp_id={company_id}'
                response = requests.get(popup_url, headers=headers)
                soup = BeautifulSoup(response.text, "lxml")
                telephone = soup.find('p', style='margin-bottom: 5px;').text.replace('Ð¢ÐµÐ»ÐµÑÐ¾Ð½: ', '').strip()
                mail = None
                site = soup.find('a').text.replace('\n', '').strip()
            except Exception:
                telephone = None
                mail = None
                site = None
        yield data, country, reputation, address, telephone, mail, site


def writer(parametr):
    book = xlsxwriter.Workbook(r"C:\ruscable.xlsx")
    page = book.add_worksheet('Компании')

    row = 0
    column = 0

    page.set_column("A:A", 30)
    page.set_column("B:B", 30)
    page.set_column("C:C", 20)
    page.set_column("D:D", 50)
    page.set_column("E:E", 20)
    page.set_column("F:F", 20)
    page.set_column("F:F", 20)

    page.write(row, column, 'Название')
    page.write(row, column + 1, 'Страна, город')
    page.write(row, column + 2, "Репутация")
    page.write(row, column + 3, "Адрес")
    page.write(row, column + 4, "Телефон")
    page.write(row, column + 5, "email")
    page.write(row, column + 6, "Сайт")
    row += 1

    for item in parametr():
        page.write(row, column, item[0])
        page.write(row, column + 1, item[1])
        page.write(row, column + 2, item[2])
        page.write(row, column + 3, item[3])
        page.write(row, column + 4, item[4])
        page.write(row, column + 5, item[5])
        page.write(row, column + 6, item[6])
        row += 1

    book.close()


writer(array)
