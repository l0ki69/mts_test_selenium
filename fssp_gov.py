from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, \
    ElementNotInteractableException, StaleElementReferenceException, TimeoutException

from captcha import read_captcha
from excel_interaction import Debtors, ProcessingExcel


class Fssp:
    """Main class for interacting  with https://fssp.gov.ru website"""

    def __init__(self):
        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 10)

        self.browser.implicitly_wait(10)
        self._restart_session()

    def __del__(self):
        self.browser.quit()

    def _restart_session(self):
        """Restarts the session if something went wrong"""

        self.browser.get('http://fssprus.ru/')

        # Закрываем всплывающее окно
        pop_up_window_button = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.tingle-modal__close')))
        pop_up_window_button.click()

        # Ждем пока станет видна кнопка расширенного поиска и кликаем найти
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn-light')))
        self.browser.find_element_by_css_selector('div.main-form__btn:nth-child(2) > button:nth-child(1)').click()

        # Ждем пока станет видно кнопку физ лицо и кликаем на нее
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.s-3:nth-child(1) > label:nth-child(1)')))
        self.browser.find_element_by_css_selector('div.s-3:nth-child(1) > label:nth-child(1)').click()

    def _introduces_captcha(self):
        """Solves captcha using Tesseract-ocr"""

        # Проверка наличия капчи на странице.
        try:
            self.browser.find_element_by_css_selector('.popup')
        except NoSuchElementException:
            # Капчи нет - мы уже авторизованы.
            return

        while True:
            try:
                # Делаем скриншот участка с капчей
                img_save = self.browser.find_element_by_id('capchaVisual').screenshot('img.png')

                if not img_save:  # Если скриншот сделать нее удалось
                    self.browser.find_element_by_css_selector('#ncapcha-submit').click()
                    continue

                text = read_captcha()
                # ВВодим капчу
                self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#captcha-popup-code')))
                self.browser.find_element_by_id('captcha-popup-code').send_keys(text)

                self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#ncapcha-submit')))
                self.browser.find_element_by_css_selector('#ncapcha-submit').click()
                # Если кнопка доступна то кликаем чтобы начать проверку корректности капчи

                # Иногда сайт может обрабатывать капчу 10+ секунд, поэтому ждем пока кнопка продолжить
                # снова станет доступной, чтобы продолжить ввод капчи
                self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#ncapcha-submit')))

            # Когда капча решается успешно окно с капчей закрывается
            except (NoSuchElementException, ElementNotInteractableException, TimeoutException):
                break
            except StaleElementReferenceException:
                continue

    def get_data(self, data_debtor: Debtors):
        """
        Unloads data about open enforcement proceedings, if any

        :param data_debtor: Information about the potential debtor

        :return: True - If you managed to upload data about the debtor
                 False - If something went wrong and it's worth starting over
        :rtype: bool
        """

        def get_text(element):
            """Gets the text of the HTML block"""
            return element.text

        # Debt list
        debts_data = list()

        for key in data_debtor.__dict__.keys():
            input_form = self.browser.find_element_by_name(f"is[{key}]")
            input_form.clear()
            input_form.send_keys(data_debtor.__getattribute__(key))

        self.browser.find_element_by_id('btn-sbm').click()

        self._introduces_captcha()

        try:
            res_table = self.browser.find_element_by_css_selector('div.results')
        except NoSuchElementException:
            # Что пошло не так, начнем сначала
            self._restart_session()
            return False

        # Если нет задолженностей
        try:
            res_table.find_element_by_css_selector('div.results-frame').find_element_by_css_selector('tbody')
        except NoSuchElementException:
            name: str = data_debtor.last_name + ' ' + data_debtor.first_name + ' ' \
                        + data_debtor.patronymic + '\n' + data_debtor.date
            debts_data.append((name, 'Нет задолженностей'))
            debts.append(debts_data)
            return True

        # Задолженности есть

        debts_info: list = self.browser.find_element_by_css_selector('div.results').find_element_by_css_selector(
            'div.results-frame').find_element_by_css_selector('tbody').find_elements_by_tag_name('td')

        # Удаляем каждый информацию из столбика Сервис
        del debts_info[5:len(debts_info):8]
        debts_info.pop(0)  # Удаляем 0-й элемент с данными о республике

        # Теперь групируем данные о каждом ОТДЕЛЬНОМ исполнительном производстве
        for index in range(0, len(debts_info) // 7):
            temp_stor: list = debts_info[index * 7:index * 7 + 7]
            debts_data.append(tuple(map(get_text, temp_stor)))

        debts.append(debts_data)
        return True


if __name__ == '__main__':
    session = Fssp()
    excel = ProcessingExcel('fssprus')

    data_debtors = excel.read_excel()
    debts = []

    counter: int = 0
    while counter != len(data_debtors):
        if session.get_data(data_debtors[counter]):
            counter += 1
        else:
            continue

    excel.write_excel(debts)
