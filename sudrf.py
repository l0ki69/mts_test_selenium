from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from excel_interaction import Debtors, ProcessingExcel
import sys
import time


class Sudrf:
    """The main class for working with https://sudrf.ru/index.php?id=300#sp"""

    def __init__(self, name_court: str):
        """
        :param name_court: the name of the court in which the search for cases will be carried out
        """
        self.court: str = name_court

        self.browser = webdriver.Firefox()
        self.wait = WebDriverWait(self.browser, 5)

        self.browser.implicitly_wait(7)
        self._restart_session()

    def __del__(self):
        self.browser.quit()

    def _restart_session(self):
        """Restarts the session if something went wrong"""

        self.browser.get('https://sudrf.ru/index.php?id=300#sp')

        try:
            self.wait.until(EC.visibility_of_element_located((By.ID, 'spSearchArea')))
            table = self.browser.find_element_by_id('spSearchArea')

            regions = Select(table.find_element_by_id('court_subj'))
            regions.select_by_visible_text('Город Москва')

            name_suds = Select(table.find_element_by_id('suds_subj'))
            name_suds.select_by_visible_text(self.court)
        except TimeoutException:
            self._restart_session()

    def get_judicial_act(self, defendant: Debtors) -> bool:
        """
        Receives all lawsuits from sudrf site
        :param defendant: Information about the potential defendant
        :return: Returns True if the person was successfully processed, otherwise False
        """

        full_name: str = ''  # В таблице а разных колонках находится ФИО, объединяем их
        for key in defendant.__dict__.keys():
            full_name += defendant.__getattribute__(key) + ' '

        try:  # Вводим ФИО
            table = self.browser.find_element_by_id('spSearchArea')

            f_name_form = table.find_element_by_id('f_name')
            f_name_form.clear()
            f_name_form.send_keys(full_name)

        except TimeoutException:
            return False

        # Нажимаем найти
        button = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[4]/div[6]/form/table/tbody/tr[8]/td[2]/input[1]')))
        button.click()
        time.sleep(5)  # Немного ждем, чтобы суд. дела прогрузились

        result_table = table.find_element_by_id('resulfs')
        result = result_table.find_elements_by_tag_name('td')  # Достаем всю таблицу с суд. делами

        data_text = list(map(lambda el: el.text, result))  # Выгружаем все данные из таблицы и конвертируем их в str

        # Удаляем титульную строку
        del data_text[0:7]

        court_cases = list()  # Хранит данные о судебных делал конкретного человека

        if not data_text: # Если список пустой, то судебные дела отсутствуют
            judicial_acts.append([[full_name, 'Судебных дела отсутствуют']])
            return True

        # Судебных дел может быть много, поэтому перебираем ячейки с данными и группируем их по 7
        for index in range(0, len(data_text) // 7):
            temp_store: list = [full_name]  # Временное хранилище для одного суд. дела
            temp_store.extend(data_text[index * 7:index * 7 + 7])  # Добавляем данные о суд. деле

            court_cases.append(temp_store)  # Добавялем дело, в общий список дел на конкретного человека

        judicial_acts.append(court_cases)  # Добавляем весь список дел на человека в общий список

        return True


if __name__ == '__main__':
    excel = ProcessingExcel('sudrf')

    data_defendant = excel.read_excel()
    if not data_defendant:
        sys.exit('Input data table does not exist')

    session = Sudrf(excel.get_court())

    judicial_acts = list()  # Stores already processed data on judicial acts

    count: int = 0
    while count != len(data_defendant):
        if session.get_judicial_act(data_defendant[count]):
            count += 1
        else:
            continue

    excel.write_excel(judicial_acts)
