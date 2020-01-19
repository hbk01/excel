import hashlib
import unittest
from collections import OrderedDict
from os import remove

import xlrd

import Excel


def get_hash(file):
    line = file.readline()
    md5 = hashlib.md5()
    while line:
        md5.update(line)
        line = file.readline()
    return md5.hexdigest()


def equal_hash(file1, file2) -> bool:
    str1 = get_hash(file1)
    str2 = get_hash(file2)
    return str1 == str2


def read_excel(file):
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)

    first_row = sheet.first_visible_rowx  # 行
    first_col = sheet.first_visible_colx  # 列

    head = []
    data = []
    last_col = 0
    max_data = 9999

    # get title data
    for col in range(first_col, first_col + max_data):
        try:
            value = sheet.cell_value(first_row, col)
            head.append(value)
        except IndexError:
            last_col = col
            break

    for row in range(first_row + 1, first_row + max_data):
        temp = []
        for col in range(first_col, last_col):
            try:
                value = sheet.cell_value(row, col)
                temp.append(value)
            except IndexError:
                break

        # temp.length = 0, it's no data
        if len(temp) == 0:
            break

        # Use OrderedDict to keep order
        dictionary = OrderedDict()
        for i in range(len(head)):
            dictionary[head[i]] = temp[i]
        data.append(dictionary)

    return data


class ExcelTest(unittest.TestCase):
    def setUp(self):
        self.src_xls = "../test/src.xls"
        self.src_txt = "../test/src.txt"
        self.src_json = "../test/src.json"
        self.temp = "../test/temp"

    def test_text2json(self):
        Excel.ExcelTools.Text2Json(self.src_txt, self.temp)
        with open(self.src_json, "r") as json, open(self.temp, "r") as text:
            data_json = json.readlines()
            data_text = text.readlines()
            self.assertEqual(data_json, data_text)
        remove(self.temp)

    def test_json2text(self):
        Excel.ExcelTools.Json2Text(self.src_json, self.temp)
        with open(self.src_txt, "r") as src, open(self.temp, "r") as out:
            data1 = src.readlines()
            data2 = out.readlines()
            self.assertEqual(data1, data2)
        remove(self.temp)

    def test_excel2json(self):
        Excel.ExcelTools.Excel2Json(self.src_xls, 0, self.temp, indent=2)
        with open(self.temp, "r") as file, open(self.src_json, "r") as src:
            data1 = file.readlines()
            data2 = src.readlines()
            self.assertEqual(data1, data2)
        remove(self.temp)

    def test_excel2text(self):
        Excel.ExcelTools.Excel2Text(self.src_xls, 0, self.temp)
        with open(self.temp, "r") as file, open(self.src_txt, "r") as src:
            data1 = file.readlines()
            data2 = src.readlines()
            self.assertEqual(data1, data2)
        remove(self.temp)

    def test_json2excel(self):
        Excel.ExcelTools.Json2Excel(self.src_json, self.temp, "sheet")
        xls_temp = read_excel(self.temp)
        xls_src = read_excel(self.src_xls)
        self.assertEqual(xls_temp, xls_src)
        remove(self.temp)

    def test_text2excel(self):
        Excel.ExcelTools.Text2Excel(self.src_txt, self.temp, "sheet")
        xls_temp = read_excel(self.temp)
        xls_src = read_excel(self.src_xls)
        self.assertEqual(xls_temp, xls_src)
        remove(self.temp)


if __name__ == '__main__':
    unittest.main()
