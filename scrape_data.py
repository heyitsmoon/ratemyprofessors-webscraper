from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import time

def scrape_data(browser, url):
    print("Collecting data... (This more reviews the professor has the longer this will take)")
    browser_options = Options() 
    browser_options.add_argument('--headless')
    if browser == "chrome":
        driver = webdriver.Chrome(options=browser_options) #makes chrome launch in the background
    if browser == "firefox":
        driver = webdriver.Firefox(options=browser_options) #makes chrome launch in the background
    driver.get(url)

    time.sleep(5)
    # try: #closes popups
    #     button_xpath = "//html/body/div[5]/div/div/button"
    #     WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath))).click()
    # except:
    #     pass
    # try:
    #     WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bx-close-inside-1177612"]'))).click()
    # except:
    #     pass
    # try:
    #     WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pendo-close-guide-f789ac0a"]'))).click()
    # except:
    #     pass

    #loads all the reviews
    while(True):
        try:                                                        
            loadmore = driver.find_element(By.XPATH, "//div[@class= 'react-tabs__tab-panel react-tabs__tab-panel--selected']/button[@class='Buttons__Button-sc-19xdot-1 PaginationButton__StyledPaginationButton-txi1dr-1 joxzkC']")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",loadmore)
            ActionChains(driver)\
                .scroll_by_amount(0, 0)\
                .perform()
            time.sleep(1)
            try:
                loadmore.click()
            except Exception as e:
                print(f"An error occurred: {e}")
                print("trying again")
                continue
            time.sleep(1)
        #once the button no longer exists, grabs the html and puts it in pagesource
        except NoSuchElementException: 
            print("Finished Collecting.")
            pagesource = driver.page_source
            driver.close()
            break
    Soup = BeautifulSoup(pagesource, "html.parser")
    Soup = BeautifulSoup(Soup.prettify(),"html.parser")
    return Soup
