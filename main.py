import time
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from utils import collect_data, create_excel


def headless_driver(proxy):
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument(f'--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36')
    options.add_argument('--disable-blink-features=AutomationControlled')
    # options.add_argument(f'--proxy-server={proxy}')
    prefs = {
            'profile.default_content_setting_values': {
                'images': 2,
                'plugins': 0,
                'popups': 2,
                'scripts': 0
            }
        }
    options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Function;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            """
    })
    
    return driver

def main_driver(proxy):
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument(f'--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36')
    options.add_argument('--disable-blink-features=AutomationControlled')
    prefs = {
            'profile.default_content_setting_values': {
                'images': 2,
                'plugins': 0,
                'popups': 2,
                'scripts': 0
            }
        }
    options.add_experimental_option('prefs', prefs)
    
    # options.add_argument(f'--proxy-server={proxy}')

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Object;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Function;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Proxy;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            """
    })
    
    return driver

  
        
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
    main_url = main_view(proxy)
    
    create_excel()
    driver = headless_driver(proxy)
    try:
        driver.get('https://movizor-info.ru/filter/')
        time.sleep(5)
        with open('cookies.pkl', 'rb') as f:
            cookies = pickle.load(f)
        for cookie in cookies:
            driver.add_cookie(cookie)
            driver.get(main_url)
        for i in range(2, int(pages_count) + 1):
            links_list = []
            links = driver.find_elements(By.CLASS_NAME, "text-decoration-underline")
            for link in links:
                links_list.append(link.get_attribute('href'))
            for link in links_list:
                driver.get(link)
                collect_data(driver.page_source)
                time.sleep(4)
            driver.get(main_url)
            next_page = driver.find_element(By.CLASS_NAME, "pagination").find_elements(
                By.TAG_NAME, "a")[-1].get_attribute('href')
            time.sleep(3)
            print(next_page[:-1])
            driver.get(f'{next_page[:-1]}{i}')
            print(driver.current_url)
            time.sleep(3)
            
        return 'Success'
    except Exception as e:
        print(e)
        return 'Error'
        
    finally:
        driver.close()
        driver.quit()


