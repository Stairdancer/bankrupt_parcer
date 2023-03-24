import json
import time
import os
import pandas as pd
import numpy as np
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException

# Создаём директорию по дате работы робота если её ещё нет:
download_path = f"D:\\work_projects\\otkr_test\\{datetime.date.today()}"
if not os.path.exists(download_path):
    os.makedirs(download_path)

#Разбираем исходный Excel-файл по колонкам разных типов:   
str_cols = ['Фамилия', 'Имя', 'Отчество', 'Номер паспорта', 'ИНН']
converter1 = {col: str for col in str_cols}
date_cols = ['Дата рождения', 'Дата выдачи', 'Время проверки ИНН']
converter2 = {col: np.datetime64 for col in date_cols}
data = pd.read_excel("bankrupts.xlsx", index_col='Siebel ID', converters=dict(converter1, **converter2))

url = "https://bankruptcy.kommersant.ru/search/index.php"

#Готовим словари, которые в будущем станут новыми столбцами датафрейма (номер статьи и время проверки в Ъ)
number_article_dict = dict()
date_last_checked = dict()

#Готовим нужные параметры для вебдрайвера, в т.ч. автоматическую печать в ПДФ без подтверждений и выбора пути
appState = {
"recentDestinations": [
    {
        "id": "Save as PDF",
        "origin": "local",
        "account": ""
    }
],
"selectedDestinationId": "Save as PDF",
"version": 2,
"download.default_directory": download_path,
"download.directory_upgrade": True
}

profile = {"printing.print_preview_sticky_settings.appState":json.dumps(appState),
            "savefile.default_directory": download_path}

chrome_options = Options()
chrome_options.add_experimental_option("prefs", profile) 
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) #без этого сорит в консоль
chrome_options.add_argument("--kiosk-printing")
chrome_options.add_argument("--disable-infobars")

#Запускаем вебдрайвер и переходим на искомый ресурс
s = Service("chromedriver_win32\chromedriver.exe")
browser = webdriver.Chrome(service=s, options=chrome_options)
browser.get(url)

#Итерируемся по записям о клиентах
for index, row in data.iterrows():
    try:
        #Ищем поле ввода ИНН
        inn_form = browser.find_element(By.NAME, "query")
        inn_form.clear()
        inn_form.send_keys(row['ИНН'])
        time.sleep(2)

        #Ищем кнопку "Найти"
        search_button_xpath = "/html/body[@class='sbk-authorized-not']/div[@class='layout']/div[@class='col_group']/div[@class='col-left']/div[@class='main-search-form container_Bankruptcy']/div[@class='search-sbk-block']/div[@class='mrg40']/div[@class='search-sbk__captcha-flex']/div[@class='captcha_submit text-left']/input[@class='hover search-sbk__btn one active']"
        search_button = browser.find_element(By.XPATH, search_button_xpath)
        search_button.click()
        time.sleep(4)

        #Проверяем, не вылетела ли "Ошибка валидации". Если вылетела - жмём Escape и повторяем
        if browser.find_element(By.ID, "searchError"):
            webdriver.ActionChains(browser).send_keys(Keys.ESCAPE).perform()
            time.sleep(4)
            search_button = browser.find_element(By.XPATH, search_button_xpath)
            search_button.click()
            time.sleep(4)
        
        #Многовато sleep'ов, но мне показалось, что на меньших значениях иногда срабатывает капча

        i = 1

        #Теперь хотим скачать все статьи первого слоя, для этого используем бесконечный 
        #цикл с выходом по NoSuchElementException
        while(True):
            article_title_xpath = f"/html/body[@class='sbk-authorized-not']/div[@class='layout']/div[@class='col_group']/div[@class='col-left']/div[@class='left-main-content']/div[@class='page-content seacrhTypeBankruptcy']/div[2]/div[@class='page-content-company'][{i}]/div[@class='text']/h2[@class='article_name']"
            print_button_xpath = f"/html/body[@class='sbk-authorized-not']/div[@class='layout']/div[@class='col_group']/div[@class='col-left']/div[@class='left-main-content']/div[@class='page-content seacrhTypeBankruptcy']/div[2]/div[@class='page-content-company'][{i}]/div[@class='text']/h2[@class='article_name']/span[3]/i[@class='fa fa-print js-print-this-article']"
            try:
                #Вытаскиваем из названия статьи номер
                article_title = browser.find_element(By.XPATH, article_title_xpath).text
                article_number = article_title.split()[2]

                #Находим и жмём кнопку печати
                print_button = browser.find_element(By.XPATH, print_button_xpath)
                print_button.click()
                time.sleep(5)

                #Сразу же переименовываем скачанный файл в соотвествии с условием
                new_name = "_".join([str(index), row['Фамилия'], row['Имя'][0], row['Отчество'][0], article_number]) + ".pdf"
                try:
                    os.rename(download_path + "\\ОБЪЯВЛЕНИЯ О НЕСОСТОЯТЕЛЬНОСТИ.pdf", 
                            download_path + "\\" + new_name)
                except FileExistsError:
                    print(f"Файл с именем {new_name} уже существует в данной папке!")
                    pass

                #В заготовленных словарях привязываем номер статьи и время проверки к Siebel_ID
                number_article_dict[index] = article_number 
                date_last_checked[index] = datetime.datetime.now()          
                i += 1
            except NoSuchElementException:
                break
    except ElementClickInterceptedException:
        print("Ой-ой, кажется мы поймали капчу!")
        break

#Обогащаем датафрейм и закидываем в Excel
data['Найдено в Ъ'] = number_article_dict
data['Время проверки в Ъ'] = date_last_checked
data.to_excel("bankrupts.xlsx")

#Закрываем вебдрайвер
browser.close()
browser.quit()




