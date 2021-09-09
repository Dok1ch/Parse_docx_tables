import datetime
import glob
import re
from docx.api import Document
from prettytable import PrettyTable


# Getting the list *.doc files
def get_files():
    files_names = []
    for docx in glob.glob("*.docx"):
        files_names.append(docx)

    return files_names


def get_data_from_docx(docx_files):
    tables_data_list = []
    for i in docx_files:
        document = Document(i)
        docx_table = document.tables[0]
        table_data = []
        keys = None
        for j, row in enumerate(docx_table.rows):
            text = (cell.text for cell in row.cells if cell.text)
            if j == 0:
                keys = tuple(text)
                continue

            row_data = dict(zip(keys, text))
            table_data.append(row_data)

        tables_data_list.append(table_data)

    return tables_data_list


# Get unsorted data from *.docx tables
def get_table_fields(tables_data, files):
    unsorted_list = []
    count_error = 0
    for _iteration in range(len(tables_data)):
        for j in range(len(tables_data[_iteration])):
            try:
                if list(tables_data[_iteration][0].keys())[3] == 'Оплата по мес.':
                    unsorted_list.append([tables_data[_iteration][j]['Оплата по мес.'].replace("\n", ""),
                                          tables_data[_iteration][j]['Гос.номер\nТС'].replace("\n", ""),
                                          files[_iteration]])
                elif list(tables_data[_iteration][0].keys())[3] == 'Оплата за год':
                    unsorted_list.append([tables_data[_iteration][j]['Оплата за год'].replace("\n", ""),
                                          tables_data[_iteration][j]['Гос.номер\nТС'].replace("\n", ""),
                                          files[_iteration]])
            except KeyError:
                count_error += 1

    return unsorted_list


# Converting date to str format
def date_conversion(date_string):
    date_string = re.sub("[^0-9.]", "", date_string).replace(' ', '')
    if len(date_string) == 8:
        date_string = datetime.datetime.strptime(date_string, '%d.%m.%y').strftime('%d.%m.%Y')
        return date_string
    elif len(date_string) == 10:
        date_string = datetime.datetime.strptime(date_string, '%d.%m.%Y').strftime('%d.%m.%Y')
        return date_string


def converting_sorting(unsorted_list):
    for i in range(len(unsorted_list)):
        unsorted_list[i][0] = date_conversion(unsorted_list[i][0])
    return sorted(unsorted_list, key=lambda x: (x[0].split('.')[::-1], x[-1]))


# Creating and filling PrettyTable
def create_table(array, fields):
    table = PrettyTable()
    table.field_names = fields
    for iteration in range(len(array)):
        table.add_row(array[iteration])

    return table


def main():
    files = get_files()
    tables_list = get_data_from_docx(files)
    unsorted_data = get_table_fields(tables_list, files)
    with open('out.txt', 'w', encoding='utf-8') as w:
        w.write(str(create_table(converting_sorting(unsorted_data), ["Оплата по", "Гос номер", "Файл"])))


if __name__ == "__main__":
    main()
