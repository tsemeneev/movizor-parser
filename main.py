import time

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from fake_useragent import UserAgent
from utils import collect_data, create_excel

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option('useAutomationExtension', False)
options.add_experimental_option('prefs', {
    'profile.managed_default_content_settings.images': 2,
    'profile.managed_default_content_settings.javascript': 2,
    'profile.managed_default_content_settings.plugins': 2,
    'profile.managed_default_content_settings.popups': 2})
options.add_argument(f'--user-agent={UserAgent.random}')
options.add_argument('--disable-blink-features=AutomationControlled')

driver = webdriver.Chrome(options=options)
driver.maximize_window()
actions = ActionChains(driver)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Function;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
        delete window.document.featurePolicy;
        
    """
})


def main(main_url, pages_count):
    driver.get(main_url)
    driver.implicitly_wait(10)

    links_list = []
    for i in range(2, pages_count + 1):
        links = driver.find_elements(By.CLASS_NAME, "text-decoration-underline")
        for link in links:
            links_list.append(link.get_attribute('href'))
        for link in links_list:
            driver.get(link)
            time.sleep(2)
            collect_data(driver.page_source)
        print('next page')
        driver.get(main_url)
        next_page = driver.find_element(By.CLASS_NAME, "pagination").find_elements(
            By.TAG_NAME, "a")[-1].get_attribute('href')
        driver.get(f'{next_page}?page=i')

    driver.close()
    driver.quit()


if __name__ == '__main__':
    create_excel('1.01')
    main('https://movizor-info.ru/catalog/01.1/', pages_count=10)
