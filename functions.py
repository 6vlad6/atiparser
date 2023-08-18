import os
import time

import openpyxl

import json

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
        "ставка", "дни", "груз", "компания"
    ]

    columns_code = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]

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
    :param data: массив данных
    :return:
    """

    workbook = openpyxl.load_workbook(pathfile)
    sheet = workbook.active
    for i in data:
        sheet.append(i)

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
        print("Не получилось войти в аккаунт")
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
        print("Не удалось прочитать файл с обработанными кодами")
        print(str(e))


def add_codes_to_file(file_path, codes):
    """
    Функция добавляет обработанный груз в файл
    :param file_path: путь к файлу
    :param code: код груза
    :return:
    """
    try:
        with open(file_path, "a") as f:
            if os.path.getsize(file_path) > 0:
                f.write("\n")
            for line in codes:
                f.write(line + "\n")
    except Exception as e:
        print("Не получилось добавить код груза в файл")
        print(str(e))


def get_load_info(driver, links, file, worked_loads_file):
    """
    Функция собирает информацию по грузу
    :param driver: драйвер
    :param links: массив ссылок на груз
    :param file: .xlsx файл для загрузки туда данных
    :param worked_loads_file: .txt файл отработанных грузов
    :return:
    """
    try:
        k = 1
        all_data = []
        all_codes = []
        for link in links:
            driver.get(link)

            company_name = "-"
            status = "-"
            ati_code = "-"
            distance = "-"
            company_hometown = "-"
            contact = "-"
            load_name = "-"
            weight_volume = "-/-"
            dimensions = "-"
            start_location = "-"
            load_date = "-"
            end_location = "-"
            price = "-"
            rate = "-"
            company_days = ""
            load_link = "-"
            company_link = "-"

            try:  # получение дистанции, габаритов, вес/объем, контактов
                json_data = driver.find_element(By.ID, "__NEXT_DATA__").get_attribute("innerHTML")
                # json_data = script_content.split('<script id="__NEXT_DATA__" type="application/json">')[1].split('</script>')[0]

                data = json.loads(json_data)

                ati_code = str(data['props']['pageProps']['load']['firmInfo']['id'])
                try:
                    firm_status = str(data['props']['pageProps']['load']['firmInfo']['firmType'])

                    if firm_status == "Грузовладелец-перевозчик":
                        status = "гр/вл-пер"
                    elif firm_status == "Грузовладелец":
                        status = "грузовл"
                except:
                    pass

                if (not check_code_in_file(worked_loads_file, ati_code)) and status != "-":

                    try:
                        distance = str(data['props']['pageProps']['load']['distance'])
                    except:
                        pass

                    try:
                        dimensions = f"{str(data['props']['pageProps']['load']['cargo']['size']['length'])}x" \
                                     f"{str(data['props']['pageProps']['load']['cargo']['size']['width'])}x" \
                                     f"{str(data['props']['pageProps']['load']['cargo']['size']['height'])}"
                    except:
                        pass

                    try:
                        weight_volume = f"{str(data['props']['pageProps']['load']['cargo']['weight'])} / {str(data['props']['pageProps']['load']['cargo']['volume'])}"
                    except:
                        pass

                    try:
                        author_name = str(data['props']['pageProps']['load']['firmInfo']['contacts'][0]['name'])
                        first_phone = str(data['props']['pageProps']['load']['firmInfo']['contacts'][0]['telephone'])
                        email = str(data['props']['pageProps']['load']['firmInfo']['contacts'][0]['email'])
                        second_phone = str(data['props']['pageProps']['load']['firmInfo']['contacts'][0]['mobile'])

                        contact = f"{str(author_name)}; {str(first_phone)}; {str(second_phone)}; {str(email)}"
                    except:
                        pass

                    try:
                        company_hometown = str(data['props']['pageProps']['load']['firmInfo']['location']['fullName'])
                    except:
                        pass

                    try:
                        company_name = str(data['props']['pageProps']['load']['firmInfo']['fullFirmName'])
                    except:
                        pass

                    try:
                        load_name = str(driver.find_element(By.CLASS_NAME, "sc-htoDjs.dPPpWm").text)
                    except:
                        pass

                    try:
                        start_location = str(driver.find_elements(By.CLASS_NAME, "locationFullName")[0].text)
                    except:
                        pass

                    try:
                        end_location = str(driver.find_elements(By.CLASS_NAME, "locationFullName")[1].text)
                    except:
                        pass

                    try:
                        load_date = str(driver.find_element(By.CLASS_NAME, "dateTime").text)
                    except:
                        pass

                    try:
                        price = str(data['props']['pageProps']['load']['payment']['sumWithoutNDS'])
                    except:
                        pass

                    try:
                        rate = str(driver.find_element(By.CLASS_NAME, "load-price-per-km").find_element(By.TAG_NAME, "span").text)
                    except:
                        pass

                    load_link = str(driver.current_url)
                    company_link = f"https://ati.su/firms/{ati_code}/info"

                    driver.get(f"https://ati.su/firms/{ati_code}/rating")

                    try:
                        time.sleep(10)
                        full_date_text = driver.find_element(By.CLASS_NAME, "green__cTf7").text
                        company_days = int(full_date_text.split("(")[-1].split(")")[0][:-3].strip())
                    except:
                        pass

                    all_codes.append(ati_code)
                    all_data.append([k, company_name, status, ati_code, distance, company_hometown, contact, load_name,
                                         weight_volume, dimensions, f'{start_location}, {load_date}', end_location, price,
                                         rate, company_days, load_link, company_link])

                    if links.index(link) % 10 == 0 or links.index(link) == len(links) -1:
                        load_to_excel(file, all_data)
                        add_codes_to_file(worked_loads_file, all_codes)
                        all_data = []

                    k += 1

                else:
                    return False

            except Exception as e:
                print(str(e))
                pass

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
    loads_on_page = []
    for link in loads_links:
        loads_on_page.append(link.get_attribute("href"))

    return loads_on_page
