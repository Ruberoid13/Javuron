# import os
import openpyxl
from tabulate import tabulate
import csv
import json


def open_excel(file_to_open):
    print('\nOpen_excel working...')
    workbuk = openpyxl.load_workbook(file_to_open)
    workshit = workbuk.active
    rows = []
    columns = []
    for row_index in range(1, 500):
        for col_index in range(1, 500):
            cell = workshit.cell(row=row_index, column=col_index).value
            if cell is None:
                break
            else:
                rows.append(cell)
        if workshit.cell(row=row_index, column=1).value is None or rows is None:
            break
        else:
            columns.append(rows)
            rows = []
    return columns


def open_csv(file_to_open):
    print('\nOpen_csv working...')
    with open(file_to_open) as csvfile:
        columns = list(csv.reader(csvfile, delimiter=';'))
        return columns


def open_json(file_to_open):
    print('\nOpen_json working...')
    tmplist = []
    result = []
    with open(file_to_open, encoding="utf-8") as jsonfile:
        parced = json.load(jsonfile)
        parced_list = parced.get(list(parced.keys())[0])
        result.append(list(parced_list[0].keys()))
        for parced_dicts in parced_list:
            for key in parced_dicts:
                tmplist.append(parced_dicts.get(key))
            result.append(tmplist)
            tmplist = []
        return result


def print_tabulated(dictionary):
    print(tabulate(dictionary, headers='firstrow', tablefmt='github'))


def main():
    x = open_excel('javu.xlsx')
    print_tabulated(x)
    c = open_csv('pizda.csv')
    print_tabulated(c)
    j = open_json('hui.json')
    print_tabulated(j)


main()
