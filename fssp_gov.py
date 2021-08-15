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
        # Ждем появления окна с капчей
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#ncapcha-submit')))
        except TimeoutException:
            self._restart_session()
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


if __name__ == '__main__':
    session = Fssp()
    excel = ProcessingExcel('fssprus')

    data_debtors = excel.read_excel()


