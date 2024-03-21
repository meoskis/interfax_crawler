from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import numpy as np
import os
import codecs

# Подключение вебдрайвера
def webdriver_launch(webdriver_dir: str, url: str):
    service = Service(
        executable_path=webdriver_dir)
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service,
                              options=options)
    driver.get(url=url)
    return driver

# Вход на спарк интерфакс
def login(driver, username, password, min_delay):
    find_login_field = driver.find_element(By.XPATH,
                                           '/html/body/header/div/div[1]/div[2]/div[2]/div[1]/form/div[1]/input[1]')
    find_login_field.send_keys(username)
    time.sleep(min_delay)
    find_password_field = driver.find_element(By.XPATH,
                                              '/html/body/header/div/div[1]/div[2]/div[2]/div[1]/form/div[1]/input[2]')
    find_password_field.send_keys(password)
    time.sleep(min_delay)
    find_password_field.send_keys(Keys.ENTER)
    return driver

# Считывание ИНН первой компании
def frst_company_search(driver, inn, min_delay):
    search_row = driver.find_element(By.XPATH,
                                     '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div/div[1]/div/div/div/div[1]/div/span/input')
    search_row.send_keys(inn)
    time.sleep(min_delay)
    search_row.send_keys(Keys.ENTER)
    return driver

# После поиска по ИНН необходимо проверить, что вообще хоть что-то нашлось
# ? - количество записей (берем из них первую)
# ? - ликвидированные компании (некликабельные, черные - без страницы)
def find_company_main_card(driver, inn):
    ammt_check = driver.find_elements(By.XPATH, '//html/body/div[@class="full-height"]/div[@class="layout-content-area"]/div/div/div[@class="content-area-body js-main-content"]/div/div/div[@class="list-layout__content"]/div/div[@class="sp-list-panel__content"]/div/div/div/div[@class="sp-list-summary__item"]')
    company = driver.find_element(By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[1]/div/div[2]/div/div[1]/div[1]/a/div/span/span[1]/span/span')
    company.click()
    driver.switch_to.window(driver.window_handles[1])
    print(f'Найдено компаний {len(ammt_check)}')

    main_card_path = f"C:\\Users\\207799\\Documents\\company_main_cards\\{str(inn)}.html"
    f = codecs.open(main_card_path, "w", "utf−8")
    main_card_html = driver.page_source
    f.write(main_card_html)
    return driver

def find_correct_left_panel(driver):
    left_panel_titles = driver.find_elements(By.XPATH, '//div[@class="card-menu-items__group-title"]')
    block_number = -100
    for block in range(1, len(left_panel_titles)+1):
        try:
            block_name = driver.find_element(By.XPATH,
                                             f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block}]/div/span[text()="Деятельность компании"]')
            block_number = block
            break
        except Exception:
            continue
    print(block_number)
    return(driver, block_number)

def find_smi_card(driver, block_number):
    find_smi_element = driver.find_elements(By.XPATH,
                                            f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button')
    element = -100
    for element in range(1, len(find_smi_element)+1):
        try:
            driver.find_element(By.XPATH,
                                f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button[{element}]/div/div/span[text()="Публикации в СМИ"]')
            element_number = element
            print(f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button[{element_number}]')
            driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button[{element_number}]').sendKeys(Keys.RETURN)
            break
        except Exception:
            continue
    print(element)
    return(driver)

def smi_card_opener(driver, inn, min_delay, max_delay):
    driver.refresh()
    time.sleep(min_delay)
    period_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div[@class="mass-media-controls-section__toolbar-container"]/div[@class="mass-media-dashboard-toolbar"]/table/tbody/tr/td[3]/div/span[2]/div')
    period_button.click()
    time.sleep(min_delay)

    different_periods = driver.find_elements(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[3]/div/span[2]/div/ul/li')
    for period in range(len(different_periods)):
        if different_periods[period].text.upper() == 'ПОСЛЕДНИЙ ГОД':
            period_number = period+1
            break
    year_period = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[3]/div/span[2]/div/ul/li[{period_number}]')
    year_period.click()

    time.sleep(max_delay)

    all_news = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-stat-section"]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/button/div')
    all_news.click()

    time.sleep(max_delay)

    find_organizations = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div/div')
    for title in range(1, len(find_organizations)+1):
        try:
            driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[{title}]/div/div/div/button/div[text()="Организации"]')
            print('-'*30)
            more_orgs = driver.find_element(By.XPATH, f'/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[{title}]/div/button[@class="sp-fake-link spoiler-button sp-facet__expand-btn spoiler-button_grey-text"]/div')
            more_orgs.click()
            time.sleep(min_delay)
        except Exception:
            continue

    main_card_path = f"C:\\Users\\207799\\Documents\\company_smi_publications\\{str(inn)}.html"
    f = codecs.open(main_card_path, "w", "utf−8")
    main_card_html = driver.page_source
    f.write(main_card_html)
    driver.close()

    return driver

def new_company_searcher(driver, new_inn, min_delay):
    driver.switch_to.window(driver.window_handles[0])
    input_line = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div/div/div[1]/div/span/input')
    input_line.send_keys(Keys.CONTROL + 'a' + Keys.DELETE)
    input_line = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div/div/div[1]/div/span/input')
    time.sleep(min_delay)
    input_line.send_keys(new_inn)
    time.sleep(min_delay)
    input_line.send_keys(Keys.ENTER)
    return driver

def smi_card_opener(driver, inn, min_delay, max_delay):
    period_button = driver.find_element(By.XPATH,
                                        '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div[@class="mass-media-controls-section__toolbar-container"]/div[@class="mass-media-dashboard-toolbar"]/table/tbody/tr/td[3]/div/span[2]/div')
    period_button.click()
    time.sleep(min_delay)

    different_periods = driver.find_elements(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[3]/div/span[2]/div/ul/li')
    for period in range(len(different_periods)):
        if different_periods[period].text.upper() == 'ПОСЛЕДНИЙ ГОД':
            period_number = period+1
            break
    year_period = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[3]/div/span[2]/div/ul/li[{period_number}]')
    year_period.click()

    time.sleep(max_delay)

    all_news = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-stat-section"]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/button/div')
    all_news.click()

    time.sleep(max_delay)

    find_organizations = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div/div')
    for title in range(1, len(find_organizations)+1):
        try:
            driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[{title}]/div/div/div/button/div[text()="Организации"]')
            print('-'*30)
            more_orgs = driver.find_element(By.XPATH, f'/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[{title}]/div/button[@class="sp-fake-link spoiler-button sp-facet__expand-btn spoiler-button_grey-text"]/div')
            more_orgs.click()
            time.sleep(min_delay)
        except Exception:
            continue

    main_card_path = f"C:\\Users\\207799\\Documents\\company_smi_publications\\{str(inn)}.html"
    f = codecs.open(main_card_path, "w", "utf−8")
    main_card_html = driver.page_source
    f.write(main_card_html)
    driver.close()

    return driver

def new_company_searcher(driver, new_inn, min_delay):
    driver.switch_to.window(driver.window_handles[0])
    input_line = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div/div/div[1]/div/span/input')
    input_line.send_keys(Keys.CONTROL + 'a' + Keys.DELETE)
    input_line = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div[1]/div/div/div/div[1]/div/span/input')
    time.sleep(min_delay)
    input_line.send_keys(new_inn)
    time.sleep(min_delay)
    input_line.send_keys(Keys.ENTER)
    return driver

url = "https://spark-interfax.ru/"
wb_path = r'C:\Users\207799\PycharmProjects\interfax_crawler\chromedriver-win64\chromedriver.exe'
username = 'FAUNIVER124' # FAUNIVER124 FAUNIVER125
password = 'kLC02oo' # kLC02oo wcKoepi
min_delay = 5
max_delay = 25

data = pd.read_excel('рейтинг крупнейших компаний россии по объему реализации продукции raex-600 final.xlsx')
data['ИНН'] = data['ИНН'].astype(str)
frst_inn = data['ИНН'].iloc[0]

driver = webdriver_launch(webdriver_dir=wb_path, url=url)
time.sleep(min_delay)

driver = login(driver=driver, username=username, password=password, min_delay=min_delay)
time.sleep(min_delay)

driver = frst_company_search(driver=driver, inn=frst_inn, min_delay=min_delay)
time.sleep(min_delay)

try:
    driver = find_company_main_card(driver=driver, inn=frst_inn)
except Exception:
    print('Ошибка в поиске компании')
time.sleep(min_delay)

try:
    driver, block_number = find_correct_left_panel(driver=driver)
except Exception:
    print('Ошибка в поиске блока "Деятельность компании"')
time.sleep(min_delay)

try:
    driver = find_smi_card(driver=driver, block_number=block_number)
except Exception:
    print('Ошибка в поиске и открытии блока "Публикации в СМИ')
time.sleep(min_delay * 3)

try:
    driver = smi_card_opener(driver=driver, inn=frst_inn, min_delay=min_delay, max_delay=max_delay)
except Exception:
    print('Ошибка в развертке блока "Публикации в СМИ"')
time.sleep(min_delay)

for inn in range(1, data['ИНН'].shape[0]):
    new_inn = data['ИНН'].iloc[inn]
    print(new_inn)
    driver = new_company_searcher(driver=driver, new_inn=new_inn, min_delay=min_delay)
    time.sleep(min_delay)

    try:
        driver = find_company_main_card(driver=driver, inn=new_inn)
    except Exception:
        print('Ошибка в поиске компании')
    time.sleep(min_delay)

    try:
        driver, block_number = find_correct_left_panel(driver=driver)
    except Exception:
        print('Ошибка в поиске блока "Деятельность компании"')
    time.sleep(min_delay)

    try:
        driver = find_smi_card(driver=driver, block_number=block_number)
    except Exception:
        print('Ошибка в поиске и открытии блока "Публикации в СМИ')
    time.sleep(min_delay * 3)

    # try:
    driver = smi_card_opener(driver=driver, inn=new_inn, min_delay=min_delay, max_delay=max_delay)
    # except Exception:
    #     print('Ошибка в развертке блока "Публикации в СМИ"')
    time.sleep(min_delay)