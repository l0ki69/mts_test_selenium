import win32com.client
from pathlib import *
import json


class Debtors:
    def __init__(self, last_name: str, first_name: str, patronymic: str, date: str):
        self.last_name = last_name  # Фамилия
        self.first_name = first_name  # Имя должника
        self.patronymic = patronymic  # Отчество
        self.date = date  # Дата рождения


class ProcessingExcel:

    def __init__(self, website: str):
        """
        Create a COM object for working with tables
        :param website: name of the site for parsing
        """

        with open('table_data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        self.name_table = data[website]['file_input']
        self.website = website
        self.excel = win32com.client.Dispatch('Excel.Application')

    def __del__(self):

        self.excel.Quit()

    def read_excel(self):
        """
        Reads data from an excel file. Works for as for fssp, as for sudrf
        :return: List with data on potential debtors
        :rtype: list[Debtors]
        """
        wb = self.excel.Workbooks.Open(Path.cwd() / self.name_table)
        sheet = wb.ActiveSheet

        num_row: int = 1  # кол-во строк
        num_column: int = 1  # кол-во столбцов
        while True:  # Цикл узнает кол-во не пустых строк и столбцов
            val_col = sheet.Cells(1, num_column).value
            val_row = sheet.Cells(num_row, 1).value

            if str(val_col) != 'None':
                num_column += 1

            if str(val_row) != 'None':
                num_row += 1

            if str(val_row) == 'None' and str(val_col) == 'None':
                num_row -= 2  # Убираем из расчетов титульную строку и строку с None
                num_column -= 1  # Убираем строку с None
                break

        # data_debtors - хранит все данные о потенциальных должниках
        data_debtors = list()

        # Над данный момент кол-во столбцов может быть только 3 или 4
        table_range: str = 'A2:'
        if num_column == 3:
            table_range += 'C' + str(num_row)
        else:
            table_range += 'D' + str(num_row)

        counter: int = 0
        data_row = list()  # Временное хранилище данных по строкам
        for cell in sheet.Range(table_range):

            if counter == 3 and num_column == 4:  # Есть ли столбец с датой рождения
                date = str(cell).split(' ')[0]
                data_row.append(date.split('-')[2] + '.' + date.split('-')[1] + '.' + date.split('-')[0])
            else:
                data_row.append(str(cell))

            counter += 1
            if counter == 4 or (counter == 3 and num_column == 3):
                if num_column == 3:
                    data_row.append('')

                data_debtors.append(Debtors(last_name=data_row[0],  # Добавляем данные в общий список данных
                                            first_name=data_row[1],
                                            patronymic=data_row[2],
                                            date=data_row[3]))
                counter = 0
                data_row.clear()

        wb.Close()

        return data_debtors
