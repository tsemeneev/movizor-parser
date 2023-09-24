import os
import re
from bs4 import BeautifulSoup
import openpyxl
from datetime import date


def create_excel():
    today = date.today().strftime("%d.%m.%y")
    path = os.path.join(os.path.dirname(__file__), f'companies_{today}.xlsx')
    if not os.path.exists(path):
        wb = openpyxl.Workbook()
        ws = wb.active
        

        ws.column_dimensions['A'].width = 27
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 40
        ws.column_dimensions['G'].width = 20
        ws.column_dimensions['H'].width = 25
        ws.column_dimensions['I'].width = 25
        ws.column_dimensions['J'].width = 20
        ws.column_dimensions['K'].width = 20
        ws.column_dimensions['L'].width = 40
        ws.column_dimensions['M'].width = 10
        ws.column_dimensions['N'].width = 20
        ws.column_dimensions['O'].width = 20
        ws.column_dimensions['P'].width = 20
        ws.column_dimensions['Q'].width = 15
        
        ws['A1'] = 'Сокращенное название'
        ws['B1'] = 'Полное название'
        ws['C1'] = 'Статус'
        ws['D1'] = 'ИНН'
        ws['E1'] = 'ОГРН'
        ws['F1'] = 'Юридический адрес'
        ws['G1'] = 'Город'
        ws['H1'] = 'Уставный капитал'
        ws['I1'] = 'Дата регистрации'
        ws['J1'] = 'ФИО Руководителя'
        ws['K1'] = 'Основной ОКВЭД'
        ws['L1'] = 'Основной ОКВЭД (название)'
        ws['M1'] = 'Количество сотрудников'
        ws['N1'] = 'Телефоны'
        ws['O1'] = 'Электронная почта'
        ws['P1'] = 'Сайты'
        ws['Q1'] = 'Доходы'

        wb.save(f'companies_{today}.xlsx')




def collect_data(source):
    soup = BeautifulSoup(source, 'lxml')
    
    left_card = soup.find(class_="col-md-7 mt-4").text
    right_card = soup.find(class_="col-md-5 mt-4")
    short_name = soup.find('h2', class_="page-title").text.strip()
    full_name = soup.find('div', class_="page-pretitle").text.strip()
    status = re.search(r'статус:\s(.*)\s', left_card).group(1)
    inn = re.search(r'КПП\s(\d*),\s(\d*)\s', left_card).group(1)
    ogrn = re.search(r'ОГРН\s(\d*)\s', left_card).group(1)
    address = re.search(r'адрес\s(.*)\s\s', left_card)
    address = '' if address is None else address.group(1)
    city = re.search(r'Город\s(.*)\s', left_card)
    city = '' if city is None else city.group(1)
    authorized_capital = re.search(r'капитал\s(.*)\s', left_card).group(1)
    registered_date = re.search(r'Дата регистрации\s(.*)\s', left_card).group(1)
    constacts = right_card.find_all('p') 
    phone_numbers = constacts[0].find_all('a')
    phone_numbers = ', '.join(['+7' + x.text for x in phone_numbers])
    websites = constacts[2].find_all('a')
    websites = ', '.join([x.text for x in websites])
    email = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', constacts[1].text)
    email = '' if email is None else email.group(0)

    manager_name = re.search(r'[Дд]иректор\s(.*)\s', left_card)
    manager_name = '' if manager_name is None else manager_name.group(1)

    revenues = re.search(r'Доходы за 2022 г.:\s(.*)\s', right_card.text)
    revenues = '' if revenues is None else revenues.group(1)

    main_type_activity = re.search(r'деятельности:\s{4}(.*\s.*)\s', left_card).group(1)
    code, activity_name = main_type_activity.split('\n')

    employees = re.search(r'Количество сотрудников в 2022 г.:\s(\d*)', right_card.text)
    employees = '' if employees is None else employees.group(1)

    

    company_data = [[short_name, full_name, status, inn, ogrn, address, city, authorized_capital, registered_date, manager_name,
                    code, activity_name, employees, phone_numbers, email, websites, revenues]]

    add_data_to_excel(company_data)


def add_data_to_excel(data):
    today = date.today().strftime("%d.%m.%y")
    workbook = openpyxl.load_workbook(f'companies_{today}.xlsx')
    sheet = workbook.active

    for row in data:
        sheet.append(row)

    workbook.save(f'companies_{today}.xlsx')
    print(f"Данные успешно добавлены в Excel файл: example.xlsx")
