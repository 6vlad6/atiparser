import os
import time

import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def create_excel_file(directory, filename):
    """
    Функция создает .xlsx файл в указанной директории с указанным названием
    :param directory: название папки, в которой будет находиться файл
    :param filename: название файла без расширения, просто текст
    :return: строка с результатом
    """
    columns = [
        "№ пункта", "название компании", "статус (грузовладелец или грузовладелец-перевозчик)", "код ати", "расстояние",
        "город откуда компания", "контакт (имя и тел)", "наименование груза", "вес, объем",
        "габариты если есть / упаковка, количество если есть", "город загрузки, дата", "город разгрузки", "цена",
        "ставка", "дни"
    ]

    columns_code = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]

    try:
        # Создание пути к файлу
        filepath = os.path.join(directory, filename+".xlsx")

        # Создание нового Excel файла
        workbook = openpyxl.Workbook()

        sheet = workbook.active

        for i in range(len(columns)):
            sheet[f'{columns_code[i]}1'] = columns[i]

        # Сохранение файла
        workbook.save(filepath)

        return filepath

    except:
        return False


def load_to_excel(pathfile, data):
    """
    Функция добавляет строчку в .xlsx-файл
    :param pathfile: путь к файлу
    :param data: строчка
    :return:
    """

    workbook = openpyxl.load_workbook(pathfile)
    sheet = workbook.active
    sheet.append(data)

    workbook.save(pathfile)

    return True


def create_driver(url):
    """
    Функция создает драйвер для парсинга
    :param url: адрес страницы
    :return: драйвер
    """

    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    # options.add_argument("--headless")  # работа без открытия браузера
    # options.add_argument(
    #     "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
    #     "Chrome/105.0.0.0 Safari/537.36")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    driver.get(url)

    return driver


def login_ati(driver, username, password):
    """
    Функция входит в аккаунт для парсинга
    :param driver: драйвер
    :param username: имя пользователя
    :param password: пароль
    :return:
    """

    # открытие поп-апа с формой входа
    try:
        login_btn = driver.find_element(By.CLASS_NAME, "header-login.header-login-button.glz-button.glz-is-small.glz-is-primary")
        login_btn.click()

        time.sleep(3)

        iframe_element = driver.find_element(By.CSS_SELECTOR, "iframe[title='Login popup']")
        driver.switch_to.frame(iframe_element)

        login_input = driver.find_element(By.ID, "login")
        login_input.send_keys(username)

        password_input = driver.find_element(By.ID, "password")
        password_input.send_keys(password)

        login_btn_popup = driver.find_element(By.ID, "action-login")
        login_btn_popup.click()

        time.sleep(10)
        driver.switch_to.default_content()

    except Exception as e:
        print(str(e))


def check_code_in_file(file_path, code):
    """
    Функция проверяет, был ли этот груз ранее обработан
    :param file_path: путь к файлу
    :param code: код груза
    :return: True - был обработан, False - не был обработан
    """
    try:
        with open(file_path, 'r') as file:
            for line in file:
                if line.strip() == code:
                    return True
        return False
    except Exception as e:
        print("ошибка в чек код ин файл")
        print(str(e))


def add_code_to_file(file_path, code):
    """
    Функция добавляет обработанный груз в файл
    :param file_path: путь к файлу
    :param code: код груза
    :return:
    """
    try:
        with open(file_path, 'a+') as file:
            file.seek(0)
            first_char = file.read(1)
            if not first_char:
                file.write(code)
            else:
                file.write('\n' + code)
    except Exception as e:
        print("ошибка в эдд код ту файл")
        print(str(e))


def get_load_info(driver, link, file, worked_loads_file):
    """
    Функция собирает информацию по грузу
    :param driver: драйвер
    :param link: ссылка на груз
    :param file: .xlsx файл для загрузки туда данных
    :param worked_loads_file: .txt файл отработанных грузов
    :return:
    """
    try:
        driver.get(link)

        code = str(driver.current_url).split("/")[-1]  # код ати
        if not check_code_in_file(worked_loads_file, code):
            name = driver.find_element(By.CLASS_NAME, "sc-htoDjs.dPPpWm").text
            start_location = driver.find_elements(By.CLASS_NAME, "locationFullName")[0].text
            end_location = driver.find_elements(By.CLASS_NAME, "locationFullName")[1].text
            load_date = driver.find_element(By.CLASS_NAME, "dateTime").text

            load_to_excel(file, ["-", "-", "-", code, "-", "-", "-", name, "-", "-", f'{start_location}, {load_date}', end_location,
                                 "-", "-", "-"])

            add_code_to_file(worked_loads_file, code)

        else:
            return False
    except Exception as e:
        print("ошибка в гет инфо")
        print(str(e))


def get_loads_on_page(driver):
    """
    Функция собирает ссылки на все грузы на странице
    :param driver: драйвер
    :return: массив ссылок
    """
    wait = WebDriverWait(driver, 15)
    loads_links = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "XfMtd.fmVZ6")))
    # loads_links = driver.find_elements(By.CLASS_NAME, "XfMtd.fmVZ6")
    loads_on_page = []
    for link in loads_links:
        loads_on_page.append(link.get_attribute("href"))

    return loads_on_page


def next_pagination_page(driver):
    """
    Функция нажимает на кнопку пагинации и открывает следующую страницу
    :param driver: драйвер
    :return: False - следующая страница недоступна, True - доступна
    """
    pagination_div = driver.find_element(By.CLASS_NAME, "SearchResults_actionsContainer__Y2RRQ.SearchResults_actionsContainer_bottom__3dbiy")

    try:
        btn = pagination_div.find_element(By.CLASS_NAME, "next_FJXnH hide_ilkOM.Pagination_nextTipClassName__tg_Y0.Pagination_tipInactive__aTL95")
        return False

    except:
        pagination_button = pagination_div.find_element(By.CLASS_NAME, "next_FJXnH.Pagination_nextTipClassName__tg_Y0")
        pagination_button.click()

        return True
