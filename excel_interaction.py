import win32com.client
from pathlib import *

excel = win32com.client.Dispatch('Excel.Application')


def read_excel(name_table):
    """
    Reads data from an excel file. Works for as for fssp, as for sudrf
    :param str name_table: Data table name
    :return: List with data on potential debtors
    :rtype: list
    """

    wb = excel.Workbooks.Open(Path.cwd() / 'excel table' / name_table)
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
            data_debtors.append(data_row[:])  # Добавляем данные в общий список данных
            counter = 0
            data_row.clear()

    wb.Close()
    excel.Quit()

    return data_debtors
