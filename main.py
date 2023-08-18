import sys
import time

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog

from functions import create_excel_file, load_to_excel, create_driver, login_ati, get_loads_on_page, get_load_info

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


base_url = "https://loads.ati.su/"


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Парсер'
        self.left = 100
        self.top = 100
        self.width = 700
        self.height = 800
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setStyleSheet("background-color: black; color: yellow; font: 10pt Arial;")
        style_file = "style.qss"
        with open(style_file, "r") as f:
            self.setStyleSheet(f.read())

        # Поле для выбора пути на ПК
        self.path_label = QLabel('Выберите папку для сохранения:', self)
        self.path_label.move(20, 20)

        self.path_edit = QLineEdit(self)
        self.path_edit.move(20, 40)
        self.path_edit.resize(250, 40)

        self.path_button = QPushButton('Обзор', self)
        self.path_button.move(280, 40)
        self.path_button.clicked.connect(self.open_folder)

        # Текстовое поле для ввода названия
        self.name_label = QLabel('Введите название конечного файла:', self)
        self.name_label.move(20, 100)

        self.name_edit = QLineEdit(self)
        self.name_edit.move(20, 130)
        self.name_edit.resize(250, 40)

        # Кнопка для запуска скрипта
        self.start_button = QPushButton('Запустить', self)
        self.start_button.move(400, 340)
        self.start_button.clicked.connect(self.start_script)
        # Текстбокс для вывода информации
        self.output_label = QLabel('Вывод:', self)
        self.output_label.move(20, 360)

        self.output_text = QTextEdit(self)
        self.output_text.move(20, 390)
        self.output_text.resize(493, 350)

        self.login_label = QLabel('Логин:', self)
        self.login_label.move(20, 180)

        self.login_edit = QLineEdit(self)
        self.login_edit.move(20, 200)
        self.login_edit.resize(250, 40)

        self.password_label = QLabel('Пароль:', self)
        self.password_label.move(20, 250)

        self.password_edit = QLineEdit(self)
        self.password_edit.move(20, 270)
        self.password_edit.resize(250, 40)

        self.file_label = QLabel('Файл с обработанными грузами:', self)
        self.file_label.move(20, 320)

        self.file_edit = QLineEdit(self)
        self.file_edit.move(20, 340)
        self.file_edit.resize(250, 40)

        self.file_button = QPushButton('Обзор', self)
        self.file_button.move(280, 340)
        self.file_button.clicked.connect(self.open_file)

        self.show()

    def open_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '', 'Все файлы (*)')
        if file_name:
            self.file_edit.setText(file_name)

    def open_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, 'Выберите папку')
        if folder_path:
            self.path_edit.setText(folder_path)

    def start_script(self):
        self.output_text.append('Работа началась')

        folder_path = self.path_edit.text()  # папка для хранения файла
        file_name = self.name_edit.text()  # название файла
        res = create_excel_file(folder_path, file_name)  # создание файла

        login = self.login_edit.text()  # логин учетной записи для парсинга
        password = self.password_edit.text()  # пароль учетной записи для парсинга
        worked_loads_file = self.file_edit.text()  # путь к файлу с обработанными грузами

        loads = []

        if res:
            self.output_text.append('Файл создан')

            driver = create_driver(base_url)

            choose_start = driver.find_element(By.CLASS_NAME, "glz-link.glz-is-medium")  # кнопка раскрытия списка
            choose_start.click()

            list_div = driver.find_element(By.CLASS_NAME, "glz-dropdown.glz-is-bottom-right")
            li_elems_len = len(list_div.find_element(By.TAG_NAME, "ul").find_elements(By.TAG_NAME, "li"))

            btn_to_close = driver.find_elements(By.CLASS_NAME, "glz-dropdown-toggle")[1]  # кнопка для закрытия списка
            btn_to_close.click()

            loads = []

            for i in range(li_elems_len-25):  # проходим для каждого региона отправки

                try:
                    time.sleep(5)
                    choose_start.click()
                    time.sleep(2)

                    # получение всех регионов отправки
                    list_div = driver.find_element(By.CLASS_NAME, "glz-dropdown.glz-is-bottom-right")
                    li_list = list_div.find_element(By.TAG_NAME, "ul").find_elements(By.TAG_NAME, "li")

                    li_list[i].click()  # выбор региона отправки

                    btn_show_loads = driver.find_element(By.CLASS_NAME,
                                "glz-button.glz-is-primary.glz-no-radius-right.button_xMZ8T.SearchButton_button__nfHt1")

                    btn_show_loads.click()  # показать все грузы по региону
                    time.sleep(5)

                    new_links = get_loads_on_page(driver)  # получение новых грузов на странице
                    loads.extend(new_links)  # добавление новых грузов в общий массив грузов

                    # while True:  # переход по страницам пагинации
                    #
                    #     try:  # поиск последней страницы пагинации - если блок есть, то и она есть
                    #         pages_count_block = driver.find_element(By.CLASS_NAME, "total_5B9k1")
                    #         pages_count = str(pages_count_block.text).split()[-1]
                    #
                    #         for page in range(2, int(pages_count)+1):
                    #             form_to_write = driver.find_element(By.CLASS_NAME, "glz-input.input_mEUzW")
                    #             input_block = form_to_write.find_element(By.TAG_NAME, "input")
                    #
                    #             text_in_field = len(str(input_block.get_attribute("value")))
                    #             for digit in range(int(str(text_in_field))+1):
                    #                 input_block.send_keys(Keys.BACKSPACE)
                    #             input_block.send_keys(page)  # вводим номер страницы
                    #             input_block.send_keys(Keys.RETURN)  # нажимаем клавишу Enter
                    #
                    #             time.sleep(5)
                    #             new_links = get_loads_on_page(driver)  # получение новых грузов на странице
                    #             loads.extend(new_links)  # добавление новых грузов в общий массив грузов
                    #
                    #     except Exception as e:
                    #         print(str(e))
                    #         break

                    driver.execute_script("window.scrollTo(0, 0);")
                    time.sleep(3)

                except Exception as e:
                    print(str(e))

            login_ati(driver, login, password)
            time.sleep(60)

            get_load_info(driver, loads, res, worked_loads_file)

        else:
            self.output_text.append('Файл не создан')
            self.output_text.append('Работа завершилась из-за ошибки при создании файла')

        self.output_text.append('Работа завершилась')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
