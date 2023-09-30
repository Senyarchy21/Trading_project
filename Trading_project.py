from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import numpy as np
import time
from win32api import GetSystemMetrics
import win32gui
import pyautogui
from PIL import Image
import pytesseract
import mss.tools
import datetime
import pyodbc
from Comments import Comment

driver = webdriver.Chrome()


class Credentials():
    email = '***'  # enter your e-mail
    password = '***'  # enter your password


class Variables():
    last_candle_x = 0
    last_candle_y = 0
    candle_distance = 0

    monitor_width = 0
    monitor_height = 0

    chart_border_left = 0
    chart_border_top = 0
    chart_border_right = 0
    chart_border_bottom = 0

    @classmethod
    def set_last_candle(cls, x: int, y: int) -> None:
        cls.last_candle_x = x
        cls.last_candle_y = y

    @classmethod
    def set_candle_distance(cls, distance: int) -> None:
        cls.candle_distance = distance

    @classmethod
    def set_monitor_resoluthion(cls, width: int, height: int) -> None:
        cls.monitor_width = width
        cls.monitor_height = height

    @classmethod
    def set_chart_border_position(cls, left: int, top: int, right: int, bottom: int) -> None:
        cls.chart_border_left = left
        cls.chart_border_top = top
        cls.chart_border_right = right
        cls.chart_border_bottom = bottom


class Open_window():
    def open_window(self) -> None:
        driver.maximize_window()
        driver.get("https://pocketoption.ru/")

    def login(self, email: str = Credentials().email, password: str = Credentials().password) -> None:
        self.email = email
        self.password = password

        elem = driver.find_element(
            By.CSS_SELECTOR, "a[class = 'button14c']")
        elem.click()

        elem_email = driver.find_element(By.NAME, "email")
        elem_password = driver.find_element(By.NAME, "password")

        elem_email.send_keys(self.email)
        elem_password.send_keys(self.password)

        time.sleep(60)


class Url():
    def page_url(self) -> str:
        return driver.current_url


class Get_parametres():
    def get_monitor_resoluthion(self) -> None:
        width = GetSystemMetrics(0)
        height = GetSystemMetrics(1)

        Variables().set_monitor_resoluthion(width, height)

    def get_chart_borders(self) -> None:
        w, h = Variables.monitor_width, Variables.monitor_height

        w_mid = int(round(w / 2, 0))
        h_crop = h - int(round(h / 6, 0))

        r, g, b = pyautogui.pixel(w_mid, h_crop)

        while ((r != 0) and (g != 0) and (b != 0)) and (h_crop < h - 1):
            h_crop += 1
            r, g, b = pyautogui.pixel(w_mid, h_crop)

        bottom = h_crop - 1

        left = 0

        r, g, b = pyautogui.pixel(left, bottom)
        r_check, g_check, b_check = pyautogui.pixel(left, bottom)

        while (r == r_check) and (g == g_check) and (b == b_check):
            left += 1
            r, g, b = pyautogui.pixel(left, bottom)

        right = w - 1

        for i in range(3):
            r_check, g_check, b_check = pyautogui.pixel(right, bottom)
            r, g, b = pyautogui.pixel(right, bottom)

            while (r == r_check) and (g == g_check) and (b == b_check):
                right -= 1
                r, g, b = pyautogui.pixel(right, bottom)

        h_crop = int(round(h / 2, 0))

        r, g, b = pyautogui.pixel(right + 1, h_crop)
        r_check, g_check, b_check = pyautogui.pixel(right + 1, h_crop)

        while (r == r_check) and (g == g_check) and (b == b_check):
            h_crop -= 1
            r, g, b = pyautogui.pixel(right + 1, h_crop)

        top = h_crop - 1

        Variables().set_chart_border_position(left, top, right, bottom)


class Custom():
    def balance_type(self, type: str) -> None:
        self.type = type

        start_url = ''

        for i, letter in enumerate(Url().page_url()):
            if (letter == '/') and (Url().page_url()[i+1] == 'r') and (Url().page_url()[i+2] == 'u'):
                break
            else:
                start_url += letter

        type_dict = {
            'real': start_url + '/ru/cabinet/quick-high-low/',
            'demo': start_url + '/ru/cabinet/demo-quick-high-low/'
        }

        driver.get(type_dict[self.type])

        time.sleep(0.5)

    def select_option(self, product: str, position: str) -> None:
        self.product = product
        self.position = position

        elem = driver.find_element(
            By.CSS_SELECTOR, "div[class = 'currencies-block']")
        elem.click()

        time.sleep(0.5)

        product_dict = {
            'curr': "//div[@class='assets-block__nav']//*[contains(text(), 'Валюты')]",
            'crypto': "//div[@class='assets-block__nav']//*[contains(text(), 'Криптовалюты')]",
            'commodities': "//div[@class='assets-block__nav']//*[contains(text(), 'Сырьевые товары')]",
            'shares': "//div[@class='assets-block__nav']//*[contains(text(), 'Акции')]",
            'index': "//div[@class='assets-block__nav']//*[contains(text(), 'Индексы')]",
            'fav': "//div[@class='assets-block__nav']//*[contains(text(), 'Избранное')]"
        }

        elem = driver.find_element(
            By.XPATH, product_dict[self.product])
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, f"//a[@class='alist__link']//*[contains(text(), '{self.position}')]")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, "//div[@class = 'drop-down-modal-wrap undefined active']")
        elem.click()

        time.sleep(0.5)

    def scale(self, type: str) -> None:
        self.type = type

        type_dict = {
            'M10': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'M10')]",
            'M30': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'M30')]",
            'H1': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'H1')]",
            'H4': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'H4')]",
            'D1': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'D1')]",
            'D7': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'D7')]",
            'D14': "//ul[@class = 'dropdown-menu inner ']//*[contains(text(), 'D14')]"
        }

        elem = driver.find_element(
            By.CSS_SELECTOR, "button[class = 'btn dropdown-toggle btn-default']")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, type_dict[self.type])
        elem.click()

        time.sleep(0.5)

    def chart_type(self, type: str) -> None:
        self.type = type

        type_dict = {
            'line': "//div[@class = 'chart-list-block']/ul[@class = 'list-links color-blue']//*[contains(text(), 'Линия')]",
            'candle': "//div[@class = 'chart-list-block']/ul[@class = 'list-links color-blue']//*[contains(text(), 'Свечи')]",
            'bar': "//div[@class = 'chart-list-block']/ul[@class = 'list-links color-blue']//*[contains(text(), 'Бары')]",
            'haikin ashi': "//div[@class = 'chart-list-block']/ul[@class = 'list-links color-blue']//*[contains(text(), 'Хейкен Аши')]"
        }

        elem = driver.find_element(
            By.CSS_SELECTOR, "a[class='tooltip2 items__link items__link--chart-type']")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, type_dict[self.type])
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, "//div[@class = 'drop-down-modal-wrap drop-down-modal-wrap-appear-done drop-down-modal-wrap-enter-done']")
        elem.click()

        time.sleep(0.5)

    def time(self, timeline: str) -> None:
        self.timeline = timeline

        elem = driver.find_element(
            By.CSS_SELECTOR, "a[class='tooltip2 items__link items__link--chart-type']")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, f"//ul[@class = 'list-links color-blue']//*[contains(text(), '{self.timeline}')]")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, "//div[@class = 'drop-down-modal-wrap drop-down-modal-wrap-appear-done drop-down-modal-wrap-enter-done']")
        elem.click()

        time.sleep(0.5)

    def narrow_chart(self) -> None:
        x, y = Variables.chart_border_right, Variables.chart_border_top

        pyautogui.moveTo(x - 10, y + 10)
        pyautogui.drag(0, 100, 1, button='left')

    def autoscroll(self) -> None:
        elem = driver.find_element(
            By.CSS_SELECTOR, "a[class='tooltip2 items__link items__link--chart-type']")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, "//div[@class = 'settings-block']//div[@class = 'settings-row']//*[contains(text(), 'Включить автопрокрутку')]")
        elem.click()

        time.sleep(0.5)

        elem = driver.find_element(
            By.XPATH, "//div[@class = 'drop-down-modal-wrap drop-down-modal-wrap-appear-done drop-down-modal-wrap-enter-done']")
        elem.click()

        time.sleep(0.5)


class Parse():
    def get_last_candle_position(self) -> None:
        w, h = Variables().monitor_width, Variables().monitor_height
        w_locator = Variables.chart_border_right
        h_locator = int(
            round((Variables.chart_border_top + Variables.chart_border_bottom) / 2, 0))
        bottom = Variables().chart_border_bottom

        pyautogui.moveTo(w_locator, h_locator)

        color_check_up_left = (122, 140, 162)
        color_check_up_right = (84, 96, 115)
        color_check_down_left = (113, 129, 150)
        color_check_down_right = (99, 114, 133)

        color_up_left = (0, 0, 0)
        color_up_right = (0, 0, 0)
        color_down_left = (0, 0, 0)
        color_down_right = (0, 0, 0)

        while (color_up_left != color_check_up_left) and (color_up_right != color_check_up_right) and (color_down_left != color_check_down_left) and (color_down_right != color_check_down_right):
            pyautogui.moveRel(-1, 0)

            color_up_left = pyautogui.pixel(113 + 1, bottom - 69 - 61)
            color_up_right = pyautogui.pixel(
                113 + 53 + 2, bottom - 69 - 61 + 1)
            color_down_left = pyautogui.pixel(113, bottom - 69)
            color_down_right = pyautogui.pixel(113 + 53, bottom - 69)

        x_r, y_r = pyautogui.position()

        while (color_up_left == color_check_up_left) or (color_up_right == color_check_up_right) or (color_down_left == color_check_down_left) or (color_down_right == color_check_down_right):
            pyautogui.moveRel(-1, 0)

            color_up_left = pyautogui.pixel(113 + 1, bottom - 69 - 61)
            color_up_right = pyautogui.pixel(
                113 + 53 + 2, bottom - 69 - 61 + 1)
            color_down_left = pyautogui.pixel(113, bottom - 69)
            color_down_right = pyautogui.pixel(113 + 53, bottom - 69)

        pyautogui.moveRel(1, 0)

        x_l, y_l = pyautogui.position()

        pyautogui.moveRel(-1, 0)

        while (color_up_left != color_check_up_left) and (color_up_right != color_check_up_right) and (color_down_left != color_check_down_left) and (color_down_right != color_check_down_right):
            pyautogui.moveRel(-1, 0)

            color_up_left = pyautogui.pixel(113 + 1, bottom - 69 - 61)
            color_up_right = pyautogui.pixel(
                113 + 53 + 2, bottom - 69 - 61 + 1)
            color_down_left = pyautogui.pixel(113, bottom - 69)
            color_down_right = pyautogui.pixel(113 + 53, bottom - 69)

        x_r_2, y_r_2 = pyautogui.position()

        Variables().set_last_candle(int(round((x_r + x_l) / 2, 0)), h_locator)
        Variables().set_candle_distance(x_r - x_r_2 + 1)

    def move_to_last_candle(self) -> None:
        pyautogui.moveTo(Variables.last_candle_x, Variables.last_candle_y)

    def get_previous_candle(self) -> None:
        pyautogui.mouseDown(Variables.last_candle_x -
                            Variables.candle_distance, Variables.last_candle_y, button='left')
        pyautogui.moveTo(Variables.last_candle_x,
                         Variables.last_candle_y, 0.3)
        pyautogui.mouseUp(button='left')

    def get_previous_candle_new(self) -> None:
        bottom = Variables().chart_border_bottom

        color_check_up_left = (122, 140, 162)
        color_check_up_right = (84, 96, 115)
        color_check_down_left = (113, 129, 150)
        color_check_down_right = (99, 114, 133)

        color_up_left = (0, 0, 0)
        color_up_right = (0, 0, 0)
        color_down_left = (0, 0, 0)
        color_down_right = (0, 0, 0)

        while (color_up_left != color_check_up_left) and (color_up_right != color_check_up_right) and (color_down_left != color_check_down_left) and (color_down_right != color_check_down_right):
            pyautogui.moveRel(-1, 0)

            color_up_left = pyautogui.pixel(113 + 1, bottom - 69 - 61)
            color_up_right = pyautogui.pixel(
                113 + 53 + 2, bottom - 69 - 61 + 1)
            color_down_left = pyautogui.pixel(113, bottom - 69)
            color_down_right = pyautogui.pixel(113 + 53, bottom - 69)

    def get_quotes(self) -> None:
        w, h = Variables.monitor_width, Variables.monitor_height

        # Открытие
        with mss.mss() as sct_open:
            monitor = {"top": h - 134, "left": 111, "width": 127, "height": 13}

            sct_open_img = sct_open.grab(monitor)

        # Закрытие
        with mss.mss() as sct_close:
            monitor = {"top": h - 116, "left": 111, "width": 127, "height": 13}

            sct_close_img = sct_close.grab(monitor)

        # Максимум
        with mss.mss() as sct_max:
            monitor = {"top": h - 98, "left": 111, "width": 127, "height": 13}

            sct_max_img = sct_max.grab(monitor)

        # Минимум
        with mss.mss() as sct_min:
            monitor = {"top": h - 80, "left": 111, "width": 127, "height": 13}

            sct_min_img = sct_min.grab(monitor)

        open_img = Image.frombytes(
            "RGB", sct_open_img.size, sct_open_img.bgra, "raw", "BGRX")
        close_img = Image.frombytes(
            "RGB", sct_close_img.size, sct_close_img.bgra, "raw", "BGRX")
        max_img = Image.frombytes(
            "RGB", sct_max_img.size, sct_max_img.bgra, "raw", "BGRX")
        min_img = Image.frombytes(
            "RGB", sct_min_img.size, sct_min_img.bgra, "raw", "BGRX")

        open_string = pytesseract.image_to_string(open_img, 'rus')
        close_string = pytesseract.image_to_string(close_img, 'rus')
        max_string = pytesseract.image_to_string(max_img, 'rus')
        min_string = pytesseract.image_to_string(min_img, 'rus')

        open_string = open_string[:-1]
        close_string = close_string[:-1]
        max_string = max_string[:-1]
        min_string = min_string[:-1]

        open_string = open_string[10:]
        if open_string[-1:] == '.':
            open_string = open_string[:-1]

        close_string = close_string[10:]
        if close_string[-1:] == '.':
            close_string = close_string[:-1]

        max_string = max_string[10:]
        if max_string[-1:] == '.':
            max_string = max_string[:-1]

        min_string = min_string[9:]
        if min_string[-1:] == '.':
            min_string = min_string[:-1]

        if '.' in open_string:
            point_position = open_string.index('.')
        elif '.' in close_string:
            point_position = close_string.index('.')
        elif '.' in max_string:
            point_position = max_string.index('.')
        elif '.' in min_string:
            point_position = min_string.index('.')
        else:
            point_position = 0

        if point_position == 0:
            pass
        elif '.' in open_string:
            pass
        else:
            open_string = open_string[:point_position] + \
                '.' + open_string[point_position:]

        if point_position == 0:
            pass
        elif '.' in close_string:
            pass
        else:
            close_string = close_string[:point_position] + \
                '.' + close_string[point_position:]

        if point_position == 0:
            pass
        elif '.' in max_string:
            pass
        else:
            max_string = max_string[:point_position] + \
                '.' + max_string[point_position:]

        if point_position == 0:
            pass
        elif '.' in min_string:
            pass
        else:
            min_string = min_string[:point_position] + \
                '.' + min_string[point_position:]

        if point_position == 0:
            open_quote = 'NULL'
            close_quote = 'NULL'
            max_quote = 'NULL'
            min_quote = 'NULL'
        else:
            try:
                open_quote = float(open_string)
            except ValueError:
                open_quote = 'NULL'

            try:
                close_quote = float(close_string)
            except ValueError:
                close_quote = 'NULL'

            try:
                max_quote = float(max_string)
            except ValueError:
                max_quote = 'NULL'

            try:
                min_quote = float(min_string)
            except ValueError:
                min_quote = 'NULL'

        now = datetime.datetime.now()

        print()
        print(now, end='. ')
        print('Открытие:', open_quote, end='; ')
        print('Закрытие:', close_quote, end='; ')
        print('Максимум:', max_quote, end='; ')
        print('Минимум:', min_quote)

        return now, open_quote, close_quote, max_quote, min_quote

    def record_to_sql(self, time: datetime, open: float, close: float, max: float, min: float) -> None:
        self.time = str(time)[:4] + str(time)[7:10] + \
            str(time)[4:7] + str(time)[10:]
        self.open = open
        self.close = close
        self.max = max
        self.min = min

        conn = pyodbc.connect('Driver={SQL Server};'
                              'Server=ARSENIY\SQLEXPRESS;'
                              'Database=Trading_project;'
                              'Trusted_Connection=yes;')
        cursor = conn.cursor()

        sql = f'''insert into [dbo].[Parse_test](
                    "Datetime"
                    ,"Open"
                    ,"Close"
                    ,"Min"
                    ,"Max"
                )
                
                values(
                    cast('{self.time[:-3]}' as datetime)
                    ,{self.open}
                    ,{self.close}
                    ,{self.min}
                    ,{self.max}
                )'''

        cursor.execute(sql)

        conn.commit()


def main():
    Comment('Открытие сайта').print_first_time()

    Open_window().open_window()
    Open_window().login()

    Comment('Получение данных о разрешении монитора и положении зоны графика на мониторе').print_time()

    Get_parametres().get_monitor_resoluthion()
    Get_parametres().get_chart_borders()

    Comment('Настройка графика').print_time()

    Custom().balance_type('demo')
    Custom().select_option('curr', 'EUR/RUB OTC')
    Custom().chart_type('candle')
    Custom().scale('D14')
    Custom().time('M1')
    Custom().autoscroll()
    Custom().narrow_chart()

    Comment('Получение координат последней свечи').print_time()

    Parse().get_last_candle_position()

    # for i in range(2):
    #     pyautogui.moveTo(Variables.last_candle_x -
    #                      Variables.candle_distance * 80, Variables.last_candle_y)
    #     pyautogui.dragTo(Variables.last_candle_x,
    #                      Variables.last_candle_y, 1, button='left')

    Comment('Парсинг данных').print_time()

    check_r, check_g, check_b = pyautogui.pixel(
        Variables.chart_border_right, Variables.chart_border_top)

    for i in range(5):

        r, g, b = pyautogui.pixel(
            Variables.chart_border_right, Variables.chart_border_top)

        if (r != check_r) or (g != check_g) or (b != check_b):
            break
        else:
            Parse().get_previous_candle()
            t, open, close, max, min = Parse().get_quotes()
            Parse().record_to_sql(t, open, close, max, min)
            # time.sleep(1)
            # pyautogui.click()

    Comment().print_final_time()


if __name__ == '__main__':
    main()
