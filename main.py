import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import datetime


url = f'https://торги-россии.рф/search?categorie_childs%5B0%5D=2&trades-type=auction&photo=1&page=1'
r = requests.get(url)
soup = BeautifulSoup(r.content, 'html.parser')

qty_pages = int(soup.find('ul', class_='pagination').find_all('li')[-2].find('a')['href'].replace(
    'https://xn----etbpba5admdlad.xn--p1ai/search?categorie_childs%5B0%5D=2&trades-type=auction&photo=1&page=', ''))


def get_content(html):
    book = openpyxl.Workbook()
    book.remove(book.active)
    sheet_1 = book.create_sheet("cars")
    sheet_1['A1'].font = Font(bold=True)
    sheet_1['B1'].font = Font(bold=True)
    sheet_1['C1'].font = Font(bold=True)
    sheet_1['D1'].font = Font(bold=True)
    sheet_1['E1'].font = Font(bold=True)

    sheet_1['A1'] = 'ID'
    sheet_1['B1'] = 'Name'
    sheet_1['C1'] = 'description'
    sheet_1['D1'] = 'Price'
    sheet_1['E1'] = 'srcUTP'
    row = 2

    for page_num in range(1, qty_pages + 1):
        url = f'https://торги-россии.рф/search?categorie_childs%5B0%5D=2&trades-type=auction&photo=1&page={page_num}'
        r = requests.get(url)

        soup = BeautifulSoup(r.content, 'html.parser')
        items = soup.find_all('div', class_='card__wrapper')

        lots = []
        for item in items:
            lots.append({
                'ID': item.find('b', class_='text-primary').get_text(),
                'Name': item.find('h3').get_text(strip=True),  # strip=True удаляет лишние пробелы
                'srcUTP': 'https://торги-россии.рф/lot/' + item.find('b', class_='text-primary').get_text(),
                'description': item.find('p', class_='card__excerpt').get_text(),
                'Price': item.find('div', class_='card__bids').get("data-current-bid")
            })
            print(*lots, sep='\n')

        for lot in lots:
            sheet_1[row][0].value = str(lot['ID'])
            sheet_1[row][1].value = str(lot['Name'])
            sheet_1[row][2].value = str(lot['description'])
            sheet_1[row][3].value = str(lot['Price'])
            sheet_1[row][4].value = str(lot['srcUTP'])
            row += 1

    today = datetime.datetime.today()
    today = today.strftime("%Y-%m-%d")
    book.save(today + " ResultCars.xlsx")


def parse():
    html = requests.get(url)
    if html.status_code == 200:
        lots = get_content(html.text)
        messagebox.showinfo('Сохранение', 'Успешно сохранено!')
    else:
        print('ERROR - status_code != 200')
        messagebox.showerror('Ошибка!', 'ERROR - status_code != 200')


def window():
    root = Tk()
    root.title('Импорт машин')
    root.geometry('500x200')
    root.resizable(0, 0)

    label = Label(root, text='Импорт машин', font='Arial 15 bold')
    label.pack()
    text = Label(root,
                 text='Парсинг сайта https://торги-россии.рф по категории \n легковой транспорт, только с наличием фото',
                 font='Arial 12')
    text.pack()

    btn = Button(root, text='Создать', font='Arial 15 bold', command=parse)
    btn.pack()

    root.mainloop()


window()