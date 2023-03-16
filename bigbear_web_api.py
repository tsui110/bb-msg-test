from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

def get_company_name(orderNo):
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument("--disable-blink-features")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        # 创建浏览器对象
        browser = webdriver.Chrome(chrome_options=chrome_options)
        browser.implicitly_wait(30)
        browser.get("https://www.baidu.com")
        time.sleep(3)
        browser.find_element(By.ID, "kw").send_keys(orderNo)
        browser.find_element(By.ID, "su").click()

        ctag = browser.find_element(By.CLASS_NAME, "op_express_delivery_footer_source")
        ret = ctag.text
        browser.quit()
        return ret
    except:
        return "获取失败"