from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException,\
                                    StaleElementReferenceException, TimeoutException, InvalidElementStateException

from captcha import read_captcha
from excel_interaction import Debtors, ProcessingExcel
import sys
import urllib.request


class Fssp:
    """Main class for interacting  with https://fssp.gov.ru website"""

    def __init__(self):
        self.browser = webdriver.Firefox()
        # Если сайт сильно лагает, рекомендация поставить wait = 10 - 15
        # Если все хорошо, то для ускорения процесса можно сделать wait = 5
        wait: int = 10
        self.wait = WebDriverWait(self.browser, wait)
        self.browser.implicitly_wait(wait)
        self._restart_session()

    def __del__(self):
        self.browser.quit()

    def _restart_session(self):
        """Restarts the session if something went wrong"""

        self.browser.get('http://fssprus.ru/')
        try:
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

        except TimeoutException:
            self._restart_session()

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
                link_img = self.browser.find_element_by_id('capchaVisual').get_attribute('src')
                img_save: tuple = urllib.request.urlretrieve(link_img)
                text = read_captcha(img_save[0])
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

    def get_data(self, data_debtor: Debtors) -> bool:
        """
        Unloads data about open enforcement proceedings, if any

        :param data_debtor: Information about the potential debtor

        :return: True - If you managed to upload data about the debtor
                 False - If something went wrong and it's worth starting over
        """

        def get_text(element):
            """Gets the text of the HTML block"""
            return element.text

        for key in data_debtor.__dict__.keys():
            try:
                self.wait.until(EC.visibility_of_element_located((By.NAME, f"is[{key}]")))
                input_form = self.browser.find_element_by_name(f"is[{key}]")
                input_form.clear()
                input_form.send_keys(data_debtor.__getattribute__(key))
            except (TimeoutException, InvalidElementStateException):
                return False

        self.browser.find_element_by_id('btn-sbm').click()

        self._introduces_captcha()

        try:
            self.browser.find_element_by_css_selector('div.results')
        except NoSuchElementException:
            # Что пошло не так, начнем сначала
            self._restart_session()
            return False

        pages: list = self._pagination_search()
        num_pages: int = len(pages) - 1
        additional_pages: bool = True  # Есть ли еще страница для обработки

        while additional_pages:

            self._introduces_captcha()

            try:
                res_table = self.browser.find_element_by_css_selector('div.results')
            except NoSuchElementException:
                # Что пошло не так, начнем сначала
                continue

            # Debt list
            debts_data = list()

            # Если нет задолженностей
            try:
                table = res_table.find_element_by_css_selector('div.results-frame')\
                    .find_element_by_css_selector('tbody')
            except NoSuchElementException:
                name: str = data_debtor.last_name + ' ' + data_debtor.first_name + ' ' \
                            + data_debtor.patronymic + '\n' + data_debtor.date
                debts_data.append([name, 'Нет задолженностей'])
                debts.append(debts_data)
                return True

            # Задолженности есть

            # Получаем данные о задолженностях и регионах
            debts_info = list(map(get_text, table.find_elements_by_tag_name('td')))
            regions = list(map(get_text, table.find_elements_by_tag_name('h3')))

            # Удаляем ячейки с названием региона
            # debts_info.remove(regions[0:len(regions)])
            for reg in regions:
                debts_info.remove(reg)

            # Удаляем элементы столбика Сервис
            del debts_info[4::8]

            # Теперь групируем данные о каждом ОТДЕЛЬНОМ исполнительном производстве
            for index in range(0, len(debts_info) // 7):
                temp_store: list = debts_info[index * 7:index * 7 + 7]

                try:
                    debts_data.append(temp_store)
                except StaleElementReferenceException:
                    continue
            debts.append(debts_data)

            if num_pages > 0:
                num_pages -= 1
                pages = self._pagination_search()
                pages[len(pages) - 1].click()

            else:
                additional_pages = False

        return True

    def _pagination_search(self) -> list:
        """
        Finds out if there are additional pages with debts

        :return: Returns the current list of pages, if there are none, then returns an empty list
        """
        # Проверка наличия доп страниц с долгами
        try:
            block_pages = self.browser.find_element_by_css_selector('div.pagination')
            pages = block_pages.find_elements_by_tag_name('a')
            return pages
        except NoSuchElementException:
            return []


if __name__ == '__main__':
    excel = ProcessingExcel('fssprus')

    data_debtors = excel.read_excel()
    if not data_debtors:
        sys.exit('Input data table does not exist')

    session = Fssp()

    debts = list()  # Stores already processed data on debtors

    counter: int = 0
    while counter != len(data_debtors):
        if session.get_data(data_debtors[counter]):
            counter += 1
        else:
            continue

    excel.write_excel(debts)
