import re
from bs4 import BeautifulSoup
import openpyxl
import pandas as pd


def create_excel(name):
    df = pd.DataFrame(columns=[
        'Название компании',
        'Статус',
        'ИНН',
        'КПП',
        'ОГРН',
        'Юридический адрес',
        'Уставный капитал',
        'Дата регистрации',
        'Руководитель ФИО',
        'Основной вид деятельности',
        'Доходы за 2022 г.',
        'Доходы за 2021 г.',
        'Доходы за 2020 г.',
        'Доходы за 2019 г.',
        'Доходы за 2018 г.',
        'Количество сотрудников в 2022 г.',
        'Количество сотрудников в 2021 г.',
        'Количество сотрудников в 2020 г.',
        'Количество сотрудников в 2019 г.'
    ])
    df.to_excel(f'{name}.xlsx', index=False)


def collect_data(source):
    soup = BeautifulSoup(source, 'lxml')
    left_card = soup.find(class_="col-md-7 mt-4").text
    right_card = soup.find(class_="col-md-5 mt-4").text
    name = re.search(r'сведения\s{2}(.*\s.*)\s', left_card).group(1)
    status = re.search(r'статус:\s(.*)\s', left_card).group(1)
    inn = re.search(r'КПП\s(\d*),\s(\d*)\s', left_card).group(1)
    kpp = re.search(r'КПП\s(\d*),\s(\d*)\s', left_card).group(1)
    ogrn = re.search(r'ОГРН\s(\d*)\s', left_card).group(1)
    address = re.search(r'адрес\s(.*)\s\s', left_card)
    address = 'Не указан' if address is None else address.group(1)
    authorized_capital = re.search(r'капитал\s(.*)\s', left_card).group(1)
    registered_date = re.search(r'Дата регистрации\s(.*)\s', left_card).group(1)
    # phone_numbers = re.search(r'Телефон:\s(\(\d{3,4}\).{6,10})\(', right_card)
    # phone_numbers = 'Не указан' if phone_numbers is None else phone_numbers.group(1)
    # email = re.search(r'почта:\s(.*)\s', right_card)
    # email = 'Не указан' if email is None else email.group(1)
    # website = re.search(r'сайт:\s(.*\.\w{2,3})\s', right_card)
    # website = 'Не указан' if website is None else website.group(1)
    manager_name = re.search(r'[Дд]иректор\s(.*)\s', left_card)
    manager_name = 'Не указан' if manager_name is None else manager_name.group(1)
    revenues_for_2022 = re.search(r'Доходы за 2022 г.:\s(.*)\s', right_card)
    revenues_for_2022 = 'Не указаны' if revenues_for_2022 is None else revenues_for_2022.group(1)
    revenues_for_2021 = re.search(r'Доходы за 2021 г.:\s(.*)\s', right_card)
    revenues_for_2021 = 'Не указаны' if revenues_for_2021 is None else revenues_for_2021.group(1)
    revenues_for_2020 = re.search(r'Доходы за 2020 г.:\s(.*)\s', right_card)
    revenues_for_2020 = 'Не указаны' if revenues_for_2020 is None else revenues_for_2020.group(1)
    revenues_for_2019 = re.search(r'Доходы за 2019 г.:\s(.*)\s', right_card)
    revenues_for_2019 = 'Не указаны' if revenues_for_2019 is None else revenues_for_2019.group(1)
    revenues_for_2018 = re.search(r'Доходы за 2018 г.:\s(.*)\s', right_card)
    revenues_for_2018 = 'Не указаны' if revenues_for_2018 is None else revenues_for_2018.group(1)
    main_type_activity = re.search(r'деятельности:\s{4}(.*\s.*)\s', left_card).group(1)
    number_of_employees_2022 = re.search(r'Количество сотрудников в 2022 г.:\s(\d*)', right_card)
    number_of_employees_2022 = 'Не указано' if number_of_employees_2022 is None else number_of_employees_2022.group(1)
    number_of_employees_2021 = re.search(r'Количество сотрудников в 2021 г.:\s(\d*)', right_card)
    number_of_employees_2021 = 'Не указано' if number_of_employees_2021 is None else number_of_employees_2021.group(1)
    number_of_employees_2020 = re.search(r'Количество сотрудников в 2020 г.:\s(\d*)', right_card)
    number_of_employees_2020 = 'Не указано' if number_of_employees_2020 is None else number_of_employees_2020.group(1)
    number_of_employees_2019 = re.search(r'Количество сотрудников в 2019 г.:\s(\d*)', right_card)
    number_of_employees_2019 = 'Не указано' if number_of_employees_2019 is None else number_of_employees_2019.group(1)

    company_data = [name, status, inn, kpp, ogrn, address, authorized_capital, registered_date, manager_name,
                    main_type_activity, revenues_for_2022, revenues_for_2021, revenues_for_2020, revenues_for_2019,
                    revenues_for_2018, number_of_employees_2022, number_of_employees_2021, number_of_employees_2020,
                    number_of_employees_2019]

    add_data_to_excel(company_data)


def add_data_to_excel(data):
    workbook = openpyxl.load_workbook('1.01.xlsx')
    sheet = workbook.active

    for row in data:
        sheet.append(row)

    workbook.save('1.01.xlsx')
    print(f"Данные успешно добавлены в Excel файл: example.xlsx")
