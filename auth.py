import pickle
import time
from drivers import headless_driver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


def authenticate(mail, password):
    try:
        driver = headless_driver('')
        driver.get('https://movizor-info.ru/client/signin/')
        time.sleep(10)
        input_mail = driver.find_element(By.XPATH, '//input[@name="email"]')
        input_mail.clear()
        time.sleep(2)
        input_mail.send_keys(mail)
        input_password = driver.find_element(By.XPATH, '//input[@name="password"]')
        input_password.clear()
        time.sleep(2)
        input_password.send_keys(password)
        time.sleep(2)
        input_password.send_keys(Keys.ENTER)
        time.sleep(2)
        with open('cookies.pkl', 'wb') as f:
            pickle.dump(driver.get_cookies(), f)
        time.sleep(2)
        
            
    except Exception as e:
        return str(e)
    
    finally:
        driver.close()
        driver.quit()
     
