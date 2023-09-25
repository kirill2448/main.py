import document as document
from Tools.scripts.dutree import display

my_log = '89653098721'
my_pass = 'sucker-gohnof-Zexna2'

regions = ['Белгородская область',
           'Брянская область',
           'Владимирская область',
           'Воронежская область',
           'Ивановская область',
           'Калужская область',
           'Костромская область',
           'Курская область',
           'Липецкая область',
           'Московская область',
           'Рязанская область',
           'Смоленская область',
           'Тверская область',
           'Тульская область',
           'Ярославская область',
           'Республика Карелия',
           'Республика Коми',
           'Ненецкий автономный округ',
           'Вологодская область',
           'Ленинградская область',
           'Мурманская область',
           'Псковская область',
           'г. Санкт-Петербург',
           'Республика Адыгея',
           'Краснодарский край',
           'Астраханская область',
           'Ростовская область',
           'Республика Дагестан',
           'Кабардино-Балкарская Республика',
           'Карачаево-Черкесская Республика',
           'Чеченская Республика',
           'Ставропольский край',
           'Республика Башкортостан',
           'Республика Мордовия',
           'Республика Татарстан',
           'Удмуртская Республика',
           'Пермский край',
           'Кировская область',
           'Оренбургская область',
           'Пензенская область',
           'Самарская область',
           'Ульяновская область',
           'Курганская область',
           'Тюменская область',
           'Ханты-Мансийский автономный округ - Югра',
           'Ямало-Ненецкий автономный округ',
           'Республика Алтай',
           'Республика Хакасия',
           'Алтайский край',
           'Красноярский край',
           'Иркутская область',
           'Новосибирская область',
           'Омская область',
           'Томская область',
           'Камчатский край',
           'Приморский край',
           'Амурская область',
           'Магаданская область',
           'Еврейская автономная область',
           'Чукотский автономный округ',
           'г. Севастополь',
           'Республика Крым',
           'г. Байконур',
           'Орловская область',
           'Тамбовская область',
           'г. Москва',
           'Архангельская область',
           'Калининградская область',
           'Новгородская область',
           'Республика Калмыкия',
           'Волгоградская область',
           'Республика Ингушетия',
           'Республика Северная Осетия-Алания',
           'Республика Марий Эл',
           'Чувашская Республика - Чувашия',
           'Нижегородская область',
           'Саратовская область',
           'Свердловская область',
           'Челябинская область',
           'Республика Тыва',
           'Кемеровская область - Кузбасс',
           'Республика Саха (Якутия)',
           'Хабаровский край',
           'Сахалинская область',
           'Республика Бурятия',
           'Забайкальский край'
           ]

projects = ['Культура',
            'Цифровая Экономика',
            'Образование',
            'Жилье',
            'Экология',
            'МСП',
            'Туризм',
            'Производительность',
            'Здравоохранение',
            'Демография',
            'БКД',
            'Наука',
            'Экспорт',
            'Атом',
            'КПМИ']

sub_projects = [
    "Культура",
    "Культурная среда",
    "Творческие люди",
    "Цифровая культура",
    "Цифровая Экономика",
    "Информационная инфраструктура",
    "Цифровое государственное управление",
    "Образование",
    "Современная школа",
    "Успех каждого ребенка",
    "Цифровая образовательная среда",
    "Социальная активность",
    "Патриотическое воспитание",
    "Молодежь России",
    "Жилье и городская среда",
#   "Жилье",
    "Комфортная городская среда",
    "Расселение аварийного жилья",
    "Экология",
    "Чистая вода",
    "Чистая страна",
    "Система обращения с ТКО",
    "Уникальные водные объекты",
    "Сохранение лесов",
    'МСП',
    "Поддержка самозанятых",
    "Предакселерация",
    "Акселерация",
    "Туризм",
    "Туристическая инфраструктура",
    "Доступность туристического продукта",
    "Совершенствование управления",
    "Производительность",
    "Системные меры (производительность)",
    "Адресная поддержка (производительность)",
    "Здравоохранение",
    "Первичная помощь",
    "Здоровое сердце",
    "Борьба с онкологией",
    "Детское здравоохранение",
    "Кадры для здравоохранения",
    "Цифровое здравоохранение",
    "Экспорт медицинских услуг",
    "Первичное звено здравоохранения",
    "Демография",
    "Финансовая поддержка семей",
    "Поддержка женщин с детьми",
    "Старшее поколение",
    "Укрепление общественного здоровья",
    "Спорт - норма жизни",
    "БКД",
    "Дорожная сеть",
    "Развитие дорожного хозяйства",
    "Безопасность на дорогах",
    "Наука",
    "Инфраструктура",
    "Экспорт",
    "Экспорт продукции АПК",
    "Системные меры (экспорт)",
    "КПМИ",
    "Европа-Западный Китай",
    "Региональные авиаперевозки"]


from pyvirtualdisplay import Display
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unittest, time, re
import os
import datetime
import pandas as pd
from openpyxl import Workbook
import itertools as it
from collections import defaultdict
import numpy as np

trhold_time = 60


# всякие функции для тыканья по кнопкам

def send_keys(driver, login_password, data):
    wait = WebDriverWait(driver, trhold_time)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, login_password)))
    element.send_keys(data)


def wait_click_by_xpath(driver, xpath):
    wait = WebDriverWait(driver, trhold_time)
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()


def wait_by_css_selector(driver, css_selector):
    wait = WebDriverWait(driver, trhold_time)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))


def wait_click_by_css(driver, css_selector):
    wait = WebDriverWait(driver, trhold_time)
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector)))
    element.click()


def wait_load_by_xp(driver, xpath):
    wait = WebDriverWait(driver, trhold_time)
    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))


def go_xp(driver, xpath):
    wait = WebDriverWait(driver, trhold_time)
    element = wait.until(EC.text_to_be_present_in_element_value(By.XPATH, xpath))
    element.click()


def create_today_folder():
    #     создает папку с сегодняшней датой
    today = str(datetime.date.today())
    path_desk = r'C:\Users\kanikin\PycharmProjects'
    path = os.path.join(path_desk, today)
    if not os.path.exists(path):
        os.mkdir(path)
    path_file = os.path.join(path, f'{today}.txt')
    return path_file


def write_to_file(path, text):
    #     записать текс в файл
    with open(path, 'a') as file:
        file.write(text + '\n')


def write_to_file1(path, text1, text2, text3):
    #     записать текс в файл
    with open(path, 'a') as file:
        file.writelines([text1, text2, text3])


wd_path = r'C:\Users\kanikin\PycharmProjects\pythonProject3\chromedriver.exe'


def get_driver(wd_path):
    #     подключение драйвера парсера
    WINDOW_SIZE = "1920,1080"
    service = Service(executable_path=r'C:\Users\kanikin\AppData\Local\Chromium Ghost\chromedriver.exe')
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": r'C:\Users',
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    driver = webdriver.Chrome(service=service, options=options)

    return driver


def login_gosusl(driver):
    #     залогинтся в гасу
    url_gasu = "https://gasu-office.roskazna.ru/index"
    driver.get(url_gasu)

    send_keys(driver, '#login', my_log)
    send_keys(driver, '#password', my_pass)

    xp = "/html/body/esia-root/div/esia-login/div/div[1]/form/div[4]/button"
    wait_click_by_xpath(driver, xp)


def connect_gasu(driver):
    #     перейти в арм аналитику
    css_path = '#header > div > div.header-top > nav.secondary-menu > ul > li:nth-child(1) > a'
    wait_by_css_selector(driver, css_path)

    driver.get("https://gasu-office.roskazna.ru/redirect/pages/smnp-arm-pfnprp-prod/?date[code]=2022-12&project["
               "id]=1&project[name]=Все%20проекты&project[code]=0&project[logo]=&project[level]=0&project["
               "subjectCode]=&subject=1#/region")


def reconnect_gasu(driver):
    #     вход в гасу без ввода логинапароля
    url_gasu = "https://gasu-office.roskazna.ru/index"
    driver.get(url_gasu)
    connect_gasu(driver)


# ---пробежаться до  Карточки---
def go(driver):
    time_wait = 1.5

    xp = '/html/body/div[1]/div/div[4]/div/div[1]/div[2]/div[2]/div/div[5]/div/div/div/div/div[1]/div/div[2]/div'
    wait_click_by_xpath(driver, xp)  # тык в стрелку культуры
    xp = '/html/body/div/div/div[4]/div/div[4]/div[2]/div[2]/div/table/tbody/tr[1]/td[3]/div/div/span'
    wait_click_by_xpath(driver, xp)  # тык в стрелку области
    xp = '/html/body/div/div/div[3]/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div/span'
    wait_click_by_xpath(driver, xp)  # тык в карточку
    #xp = '/html/body/div/div/div[3]/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div/div'
    #wait_click_by_xpath(driver, xp)  # тык в РП
    #time.sleep(time_wait)


def run_projects(driver, regions, sub_projects):
    time_wait = 1.5

    path = create_today_folder()
    write_to_file(path, "from projects")

    for pr in regions:

        el = driver.find_element(By.XPATH, '/html/body')
        el.send_keys(Keys.PAGE_UP)

        time.sleep(time_wait)

       # xp = '/html/body/div[1]/div/div[1]/div[3]/div[1]/div/p'
       # wait_click_by_xpath(driver, xp)
       # xp = '/html/body/div[1]/div/div[1]/div[3]/div[2]/div[3]/div[2]'
       # wait_click_by_xpath(driver, xp)
        #xp = '/html/body/div[1]/div/div[1]/div[3]/div[2]/button'
        xp = '/html/body/div/div/div[2]/div/div[1]/div[1]/div/div/div/button/div'  # drop window
        wait_click_by_xpath(driver, xp)

        xp = '/html/body/div/div/div[2]/div/div[1]/div[1]/div/div/input'  # search
        wait_load_by_xp(driver, xp)
        el = driver.find_element(By.XPATH, xp)

        el.send_keys(Keys.BACKSPACE)
        for _ in range(1):
            el.send_keys(Keys.BACKSPACE)  # dellete text in search window

        el.send_keys(pr)

        wait_load_by_xp(driver, xp)
        el.send_keys(Keys.RETURN)
        el.send_keys(Keys.RETURN)
        el.send_keys(Keys.ARROW_DOWN)
        el.send_keys(Keys.RETURN)
        time.sleep(time_wait)

        print(pr)
        write_to_file(path, pr)

        results[pr] = ['',0,0,0,0]
        results[pr].insert(0,pr)
        results[pr].insert(1, pr)
        print('-----------')

        for dr in sub_projects:
            try:
                time.sleep(time_wait)

                el = driver.find_element(By.XPATH, '/html/body')
                el.send_keys(Keys.PAGE_UP)

                xp = '/html/body/div/div/div[3]/div/div[1]/div/div[2]/div[1]/div/div/div[2]'
                wait_click_by_xpath(driver, xp)

                xp = '/html/body/div/div/div[2]/div/div[2]/div[1]/div/div/div'  # drop window
                wait_click_by_xpath(driver, xp)

                xp = '/html/body/div/div/div[2]/div/div[2]/div[1]/div/div/input'
                time.sleep(2)
                # xp='MuiOutlinedInput-input MuiInputBase-input MuiInputBase-inputAdornedEnd MuiAutocomplete-input MuiAutocomplete-inputFocused css-1uvydh2'
                # el=driver.find_element(By.CSS_SELECTOR,xp.__getattribute__("aria-expanded"))
                # print(el)

                if dr == 'Жилье':
                    print(dr)
                    results[dr] = ['', '',0, 0, 0]
                    xp = '/html/body/div/div/section/a[3]'
                    wait_click_by_xpath(driver, xp)
                    xp = '/html/body/div/div/div[2]/div/div[2]/div[1]/div/div/input'
                    wait_click_by_xpath(driver, xp)
                    time.sleep(4)
                    xp = '/html/body/div/div/div[3]/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div/div/span'  # drop window
                    wait_click_by_xpath(driver, xp)

                    time.sleep(4)

                    xp = '/html/body/div[1]/div/div[3]/div/div[1]/div/div[2]/div[1]/div/div/div[2]/button'
                    wait_click_by_xpath(driver, xp)

                    xp = '//li[contains(.,"тыс. ₽")]'
                    wait_click_by_xpath(driver, xp)

                    write_to_file(path, dr)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]/div/p/span'
                    wait_load_by_xp(driver, xp)
                    write_to_file(path, dr)

                    results[dr].insert(0,pr)
                    results[dr].insert(1,dr)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'касса = {el.text}')  # касса
                    write_to_file(path, f'касса = {el.text}')
                    results[dr].insert(2, el.text)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[2]/div/span'
                    wait_load_by_xp(driver, xp)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'план = {el.text}')  # план
                    write_to_file(path, f'план = {el.text}')
                    results[dr].insert(3,el.text)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[3]/div/span'
                    wait_load_by_xp(driver, xp)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'факт = {el.text}')  # факт
                    write_to_file(path, f'факт = {el.text}')
                    results[dr].insert(4, el.text)
                if dr == 'Акселерация':
                    print(dr)
                    xp = '/html/body/div/div/section/a[3]'
                    wait_click_by_xpath(driver, xp)
                    xp = '/html/body/div/div/div[3]/div/div[4]/div/div[2]/div/table/tbody/tr[7]/td[1]/div/div/div[1]/span'  # drop window
                    wait_click_by_xpath(driver, xp)

                    time.sleep(2)

                    xp = '/html/body/div/div/div[2]/div/div[2]/div[1]/div/div/input'
                    wait_load_by_xp(driver, xp)
                    el = driver.find_element(By.XPATH, xp)
                    el.send_keys(Keys.PAGE_DOWN)

                    xp = '//*[@id="root"]/div/div[3]/div/div[4]/div/div[2]/div/table/tbody/tr[2]/td[1]/div/div/span'  # drop window
                    wait_click_by_xpath(driver, xp)

                    time.sleep(2)

                    el = driver.find_element(By.XPATH, '/html/body')
                    el.send_keys(Keys.PAGE_UP)
                    time.sleep(2)
                    xp = '/html/body/div[1]/div/div[3]/div/div[1]/div/div[2]/div[1]/div/div/div[2]/button'
                    wait_click_by_xpath(driver, xp)

                    xp = '//li[contains(.,"тыс. ₽")]'
                    wait_click_by_xpath(driver, xp)
                    write_to_file(path, dr)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]/div/p/span'
                    wait_load_by_xp(driver, xp)
                    write_to_file(path, dr)
                    results[dr]=['','',0,0,0]
                    results[dr].insert(0, pr)
                    results[dr].insert(1, dr)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'касса = {el.text}')  # касса
                    write_to_file(path, f'касса = {el.text}')
                    results[dr].insert(2, el.text)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[2]/div/span'
                    wait_load_by_xp(driver, xp)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'план = {el.text}')  # план
                    write_to_file(path, f'план = {el.text}')
                    results[dr].insert(3, el.text)

                    xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[3]/div/span'
                    wait_load_by_xp(driver, xp)
                    el = driver.find_element(By.XPATH, xp)
                    print(f'факт = {el.text}')  # факт
                    write_to_file(path, f'факт = {el.text}')
                    results[dr].insert(4, el.text)
                wait_load_by_xp(driver, xp)

                el = driver.find_element(By.XPATH, xp)

                for _ in range(1):
                    el.send_keys(Keys.BACKSPACE)  # dellete text in search window

                el.send_keys(dr)

                wait_load_by_xp(driver, xp)
                el.send_keys(Keys.RETURN)
                el.send_keys(Keys.RETURN)
                el.send_keys(Keys.ARROW_DOWN)
                el.send_keys(Keys.RETURN)

                time.sleep(time_wait)
                print(dr)

                write_to_file(path, dr)

                results[dr]=['','',0,0,0]
                results[dr].insert(0,pr)
                results[dr].insert(1, dr)
                xp = '/html/body/div[1]/div/div[3]/div/div[1]/div/div[2]/div[1]/div/div/div[2]/button'
                wait_click_by_xpath(driver, xp)

                xp = '//li[contains(.,"тыс. ₽")]'
                wait_click_by_xpath(driver, xp)

                xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]/div/p/span'
                wait_load_by_xp(driver, xp)
                el = driver.find_element(By.XPATH, xp)
                print(f'касса = {el.text}')  # касса
                write_to_file(path, f'касса = {el.text}')
                results[dr].insert(2,el.text)

                xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[2]/div/span'
                wait_load_by_xp(driver, xp)
                el = driver.find_element(By.XPATH, xp)
                print(f'план = {el.text}')  # план
                write_to_file(path, f'план = {el.text}')
                results[dr].insert(3,el.text)

                xp = '//*[@id="root"]/div/div[3]/div/div[1]/div/div[2]/div[2]/div/div[3]/div/span'
                wait_load_by_xp(driver, xp)
                el = driver.find_element(By.XPATH, xp)
                print(f'факт = {el.text}')  # факт
                write_to_file(path, f'факт = {el.text}')
                results[dr].insert(4,el.text)
                print(results[dr])
                print('-----------')

            except Exception as e: \
                    print(e)
            time.sleep(time_wait)
            print('-----------')


def get_table_fin():
    #     скачать таблицу финансирования
    xp = '//*[@id="root"]/div/div[4]/div/div[2]/div[2]/div/div[1]/div/div/div[3]'
    wait_click_by_xpath(driver, xp)  # тык в стрелку financing

    xp = '/html/body/div/div/div[4]/div/div[3]/div/div/div/div[2]/button'
    wait_click_by_xpath(driver, xp)  # download table

    xp = '/html/body/div[2]/div[3]/ul/li[2]'
    wait_click_by_xpath(driver, xp)  # xlsx


def save_table(result):
    today = str(datetime.date.today())
    print(results)
    columns = ['Нам','Наименование области', 'Наименование ФП', 'касса', 'план', 'факт']
    df = pd.DataFrame(columns=columns)

    for key in result:
        arr = [key]
        # print(arr)
        for value in result[key]:
            arr.append(value)
            print(arr)
        df = df._append(dict(zip(df.columns, arr)), ignore_index=True)

    df.to_excel(f"{today+' '+str(result.keys())[12:39]}.xlsx")


def rename_and_move_my_table():
    #     переименую и перенесу собственную таблицу в папку
    today = str(datetime.date.today()) + '.xlsx'
    path_desk = r'C:\Users\kanikin\PycharmProjects'
    path = os.path.join(path_desk, today)
    new_name = today
    os.rename(path, os.path.join(os.path.dirname(path), new_name))

    today = str(datetime.date.today())
    path_desk = r'C:\Users\kanikin\PycharmProjects'
    destination_folder = os.path.join(path_desk, today)
    # move the file to the destination folder
    os.replace(os.path.join(os.path.dirname(path), new_name), os.path.join(destination_folder, new_name))
    return os.path.join(destination_folder, new_name)


def rename_dowloaded_table():
    #     переименовать скачанную таблицу
    res = []
    folder_path = r'C:\Users\kanikin\PycharmProjects'

    # specify the new base name for the files
    new_base_name = 'fin'

    # get a list of all the files in the folder
    files = os.listdir(folder_path)

    # loop through each file and rename it
    for index, file_name in enumerate(files):
        # construct the new file name with a unique index
        new_file_name = f"{new_base_name}_{index + 1}.xlsx"
        # construct the full path to the original and new file names
        original_path = os.path.join(folder_path, file_name)
        new_path = os.path.join(folder_path, new_file_name)
        # rename the file
        os.rename(original_path, new_path)
        res.append(new_path)

    return res[0]


def rename_cols(data):
    cols = data.columns
    my_names = ['name', 'plan', 'fin_0', 'fin_1', 'fin_2', 'fin_3', 'kassa', 'fin_4']
    d_names = {}
    i = 0
    for col in cols:
        d_names[col] = my_names[i]
        i += 1
    data.rename(columns=d_names, inplace=True)


def proccess_fintable(path):
    data = pd.read_excel(path)
    rename_cols(data)
    data.dropna(inplace=True)
    return data


def parse_fin_table_pipeline():
    get_table_fin()  # download
    time.sleep(5)
    path = rename_dowloaded_table()  # rname it
    data = proccess_fintable(path)  # make it lookgood
    display(data.head(30))


results = {}

driver = get_driver(wd_path)

login_gosusl(driver)
connect_gasu(driver)
go(driver)
#run_projects(driver, regions[0:1], sub_projects[0:2])
for i in range(0,27,1):
    run_projects(driver, regions[0+i:1+i], sub_projects[0:60])
    save_table(results)
    results = {}

#run_projects(driver, regions[0:1], sub_projects[0:2])
#save_table(results)
#results = {}




run_projects(driver, regions[1:2], sub_projects[0:2])
save_table(results)

reconnect_gasu(driver)


# rename_and_move_my_table()

reconnect_gasu(driver)
# parse_fin_table_pipeline()

















