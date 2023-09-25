import time
import pickle
from bs4 import BeautifulSoup

from drivers import headless_driver, main_driver
from utils import data_to_excel
from selenium.webdriver.common.by import By


def main_view(proxy):
    driver = main_driver(proxy)
    driver.get('https://movizor-info.ru/filter/')
    time.sleep(5)
    with open('cookies.pkl', 'rb') as f:
        cookies = pickle.load(f)
    for cookie in cookies:
        driver.add_cookie(cookie)
    time.sleep(10)
    flag = True
    while flag:
        if 'https://movizor-info.ru/filter/view/' in driver.current_url:
            flag = False
        time.sleep(2)

    current_url = driver.current_url
    driver.close()
    driver.quit()
    return current_url


def parser(proxy, pages_count):
    print(proxy, pages_count)
    main_url = main_view(proxy)

    data_to_excel.create_excel()
    driver = headless_driver(proxy)
    try:
        driver.get(main_url)
        time.sleep(2)
        with open('cookies.pkl', 'rb') as f:
            cookies = pickle.load(f)
        for cookie in cookies:
            driver.add_cookie(cookie)
            driver.refresh()
        for i in range(2, int(pages_count) + 1):
            card_list = []
            soup = BeautifulSoup(driver.page_source, 'lxml')
            cards = soup.find("div", attrs={"class": "card-body"}).find_all("p")

            for card in cards:
                if card.text:
                    card_list.append(card.text)
            for card in card_list:
                data_to_excel.collect_data(card, proxy)
                
            driver.get(main_url)
            next_page = driver.find_element(By.CLASS_NAME, "pagination").find_elements(
                By.TAG_NAME, "a")[-1].get_attribute('href')
            time.sleep(2)
            print(next_page[:-1])
            driver.get(f'{next_page[:-1]}{i}')
            print(driver.current_url)
            time.sleep(1)

        return 'Success'
    except Exception as e:
        print(e)
        return e
    
    finally:
        driver.close()
        driver.quit()
