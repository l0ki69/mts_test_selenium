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
            self.data_json = json.load(json_file)

        self.name_table = self.data_json[website]['file_input']
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
        # Проверка существует ли табилца
        if Path(Path.cwd() / self.name_table).is_file():
            # Да, подключаемся к ней
            wb = self.excel.Workbooks.Open(Path.cwd() / self.name_table)
        else:
            return []

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
            table_range += 'C' + str(num_row + 1)
        else:
            table_range += 'D' + str(num_row + 1)

        counter: int = 0
        data_row = list()  # Временное хранилище данных по строкам
        for cell in sheet.Range(table_range):

            if counter == 3 and num_column == 4:  # Есть ли столбец с датой рождения
                if str(cell) == 'None':
                    data_row.append('')
                else:
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

    def write_excel(self, data):
        """
        Writes data to an Excel file.
        If file with such filename doesn't exist - creates it.
        Otherwise, it will overwrite the existing
        Works for as for fssp, as for sudrf

        :param data: data to write
        :type data: list[list[str]]
        """
        name_table = self.data_json[self.website]['file_output']
        pop_up_window = str(self.data_json[self.website]['pop-up_window'])

        # Проверка существует ли табилца
        if Path(Path.cwd() / name_table).is_file():
            # Да, подключаемся к ней
            wb = self.excel.Workbooks.Open(Path.cwd() / name_table)
        else:
            # Нет, создаем новую
            wb = self.excel.Workbooks.Add()
            wb.SaveAs(str(Path.cwd()) + '/' + name_table)

        # Подключаемся к активному листу и очищаем его от старых данных
        sheet = wb.ActiveSheet
        sheet.UsedRange.Delete()

        # Узнаем кол-во столбцов по титульнику
        num_columns: int = len(self.data_json[self.website]['headers'])

        # Заполняем Титульную строку
        sheet.Range('A1:' + chr(64 + num_columns) + '1').Value = self.data_json[self.website]['headers']

        # Сдвиг по таблице
        shift_cell: int = 2
        for element in data:

            if self.website == 'fssprus':
                # Иногда добавляется текс из всплывающего окна, которое появляется при наведении на поле Должник
                # Поэтому удаляем его
                for elem in element:
                    elem[0] = elem[0].replace(pop_up_window, '')

            if len(element[0]) > 2:
                # Долги есть
                sheet.Range('A' + str(shift_cell) + ':' + chr(64 + num_columns)
                            + str(shift_cell + len(element) - 1)).Value = element
                shift_cell += len(element)
            else:
                # Нет долгов
                sheet.Range('A' + str(shift_cell) + ':B' + str(shift_cell)).Value = element
                shift_cell += 1

        wb.Save()
        wb.Close()
