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


url = "https://spark-interfax.ru/"
wb_path = r'C:\Users\nkmeo\PycharmProjects\interfax_crawler\chromedriver-win64'
username = 'FAUNIVER124' # FAUNIVER124 FAUNIVER125
password = 'kLC02oo' # kLC02oo wcKoepi

data = pd.read_excel('рейтинг крупнейших компаний россии по объему реализации продукции raex-600 final.xlsx')

# Подключение вебдрайвера
service = Service(executable_path=r"c:\users\207799\pycharmprojects\pythonproject\chromedriver-win64\chromedriver.exe")
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.get(url=url)

time.sleep(5)

find_login_field = driver.find_element(By.XPATH, '/html/body/header/div/div[1]/div[2]/div[2]/div[1]/form/div[1]/input[1]')
find_login_field.send_keys(username)
time.sleep(5)
find_password_field = driver.find_element(By.XPATH, '/html/body/header/div/div[1]/div[2]/div[2]/div[1]/form/div[1]/input[2]')
find_password_field.send_keys(password)
time.sleep(5)
find_password_field.send_keys(Keys.ENTER)

time.sleep(5)

search_row = driver.find_element(By.XPATH, '/html/body/div[1]/div[5]/div[1]/div[1]/div/div/div/div[1]/div/div/div/div[1]/div/span/input')
search_row.send_keys(7730569353)
time.sleep(5)
search_row.send_keys(Keys.ENTER)

time.sleep(5)

# ------ #
ammt_check = driver.find_elements(By.XPATH, '//html/body/div[@class="full-height"]/div[@class="layout-content-area"]/div/div/div[@class="content-area-body js-main-content"]/div/div/div[@class="list-layout__content"]/div/div[@class="sp-list-panel__content"]/div/div/div/div[@class="sp-list-summary__item"]')
# print(ammt_check)
if len(ammt_check) >= 1:
    company = driver.find_element(By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div/div[1]/div/div[2]/div/div[1]/div[1]/a/div/span/span[1]/span/span')
    company.click()
    driver.switch_to.window(driver.window_handles[1])

    time.sleep(5)

    main_card_path = f"C:\\Users\\207799\\Documents\\company_main_cards\\{str(2127309097)}.html"
    #open file in write mode with encoding
    f = codecs.open(main_card_path, "w", "utf−8")
    #obtain page source
    main_card_html = driver.page_source
    #write page source content to file
    f.write(main_card_html)

    time.sleep(5)

    #----------- #
    left_panel_titles = driver.find_elements(By.XPATH, '//div[@class="card-menu-items__group-title"]')
    for block in range(len(left_panel_titles)):
        block_name = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block+1}]').text
        if 'ПУБЛИКАЦИИ В СМИ' in block_name.upper():
            block_number = block+1
            break
    # print(block_number)

    find_smi_element = driver.find_elements(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button')
    for element in range(len(find_smi_element)):
        if find_smi_element[element].text.upper() == 'ПУБЛИКАЦИИ В СМИ':
            element_number = element+1
            break
    # print(element_number)
    smi_info = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[3]/div[2]/div/div[1]/div/div[{block_number}]/button[{element_number}]')
    smi_info.click()

    time.sleep(10)

    period_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[2]/div/span[2]/div/div')
    period_button.click()

    different_periods = driver.find_elements(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[2]/div/span[2]/div/ul/li')
    for period in range(len(different_periods)):
        if different_periods[period].text.upper() == 'ПОСЛЕДНИЙ ГОД':
            period_number = period+1
            break

    year_period = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-controls-section"]/tbody/tr/td/div/div/table/tbody/tr/td[2]/div/span[2]/div/ul/li[{period_number}]')
    year_period.click()

    time.sleep(30)

    all_news = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[2]/table[@class="sp-card-content-section sp-card-content-section_width_unlimited mass-media-wide-section mass-media-stat-section"]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/button/div')
    all_news.click()

    time.sleep(20)

    find_organizations = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div/div')
    for title in range(len(find_organizations)):
        try:
            title_name = driver.find_element(By.XPATH, f'/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[@class="sidebar-filters__section sidebar-filters__section_no-delimiter"]/div/div/div/button/div[text()="Организации"]')
            more_orgs = driver.find_element(By.XPATH, f'/html/body/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[3]/div/button[6]/div')
        except Exception:
            continue
    print(title_number)

    more_organizations = driver.find_element(By.XPATH, f'/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/div/div/div/div/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[{title_number}]/div/button[@class="sp-fake-link spoiler-button sp-facet__expand-btn spoiler-button_grey-text"]/div')
    more_organizations.click()

    time.sleep(5)

    main_card_path = f"C:\\Users\\207799\\Documents\\company_smi_publications\\{str(2127309097)}.html"
    #open file in write mode with encoding
    f = codecs.open(main_card_path, "w", "utf−8")
    #obtain page source
    main_card_html = driver.page_source
    #write page source content to file
    f.write(main_card_html)

    time.sleep(200)
