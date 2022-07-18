import time
import common as com
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def set_chrome_driver():
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('--headless')
    # chrome_options.add_argument('--window-size=1920x1080')
    # chrome_options.add_argument("--disable-gpu")
    # chrome_options.add_argument('--disable-dev-shm-usage')

    # chrome_options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")
    # chrome_options.add_argument("lang=ko_KR")

    # print(com.get_desktop_path())


    # download_directory가 없는 폴더면 만들어주고, Desktop 경로는 지정 못함
    download_directory = com.get_user_profile_path() + '\\Downloads\\selenium_download'
    prefs = {'download.default_directory': download_directory}
    chrome_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'plugins', {get: function() {return[1, 2, 3, 4, 5];},});")

    return driver


if __name__ == '__main__':
    browser = set_chrome_driver()
    browser.get("http://python.org")

    browser.find_element(By.ID, 'downloads').click()
    browser.implicitly_wait(5)
    # time.sleep(2)

    element = browser.find_element(By.CSS_SELECTOR, '#content > div > section > div.row.download-list-widget > ol > li:nth-child(1) > span.release-number > a')
    element.click()
    browser.implicitly_wait(5)
    # time.sleep(2)

    element = browser.find_element(By.CSS_SELECTOR, '#content > div > section > article > table > tbody > tr:nth-child(1) > td:nth-child(1) > a')
    element.click()
    browser.implicitly_wait(5)
    # element.send_keys(Keys.ENTER)

    # time.sleep(1)
    # browser.quit()
