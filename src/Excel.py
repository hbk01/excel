#! /usr/bin/env python3
# -*- coding: utf-8 -*-

__author__ = "hbk"
__email__ = "3243430237@qq.com"
__github__ = "https://github.com/hbk01/"

from sys import argv


class ExcelTools:
    """ Excel tools by hbk """

    @staticmethod
    def Excel2Json(excel_file, sheet_index_or_name, json_file=None):
        """
        Excel to Json.
        :param excel_file: excel file path.
        :param sheet_index_or_name: what's the sheet
        :param json_file: default is None, if set it's, the json will write to this file.
        :return: json string
        """
        # import package
        from collections import OrderedDict
        import xlrd
        import json

        workbook = xlrd.open_workbook(excel_file)
        try:
            sheet_index = int(sheet_index_or_name)
            sheet = workbook.sheet_by_index(sheet_index)
        except ValueError:
            sheet = workbook.sheet_by_name(sheet_index_or_name)

        first_row = sheet.first_visible_rowx  # 行
        first_col = sheet.first_visible_colx  # 列

        print("\nfirst cell [" + str(first_row) + ", " + str(first_col) + "]")

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

        print("\nhead = " + str(head))

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
            print("find data " + str(temp))
            # Use OrderedDict to keep order
            dictionary = OrderedDict()
            for i in range(len(head)):
                dictionary[head[i]] = temp[i]
            data.append(dictionary)

        if json_file is not None:
            json.dump(data, open(json_file, "w"))

        json_string = json.dumps(data)
        return json_string  # type: str

    @staticmethod
    def Json2Excel(json_file, excel_file, sheet_name):
        """
        Json to Excel.
        :param json_file: json file.
        :param excel_file: output excel file.
        :param sheet_name: output excel sheet name.
        :return: no return
        """
        # Use OrderedDict to keep order

        from collections import OrderedDict
        import xlwt
        import json

        json_array = json.load(open(json_file, 'r'), object_pairs_hook=OrderedDict)
        print("\nLoad Json: " + str(json_array))

        workbook = xlwt.Workbook()
        workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
        sheet = workbook.get_sheet(0)

        # 写入首行
        for i, key in enumerate(json_array[0].keys()):
            sheet.write(0, i, key)

        row = 1  # 行，首行是标题
        col = 0  # 列
        for index in json_array:
            keys = index.keys()
            print("\nItem " + str(row))
            for key in keys:
                print(str(key) + ": " + str(index.get(key)))
                sheet.write(row, col, index.get(key))
                col += 1
            row += 1
            col = 0

        workbook.save(excel_file)

    @staticmethod
    def Text2Excel(text_file, excel_file, sheet_name, item_separator=" ", text_file_encoding="utf-8"):
        """
        Text to Excel
        :param text_file: text file path
        :param excel_file: output excel file path
        :param sheet_name: sheet name for the excel
        :param item_separator: characters used to separate each item in text file, it's default to a space
        :param text_file_encoding: text file encoding, it's default to utf-8
        :return: none, it's write to excel file
        """
        import xlwt

        data = []
        with open(text_file, "r", encoding=text_file_encoding) as file:
            contents = file.readlines()

            for content in contents:
                # create a function to delete '\n'
                fun = lambda x: x.replace("\n", "")
                line = list(map(fun, content.split(item_separator)))
                data.append(line)
                print("Load Text: " + str(line))

            workbook = xlwt.Workbook()
            workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            sheet = workbook.get_sheet(0)

            # 写入首行
            for i, item in enumerate(data[0]):
                sheet.write(0, i, item)
                print("%d/%d : %s" % (0, i, item))

            for row, item in enumerate(data[1:]):
                for col, text in enumerate(item):
                    print("%d/%d : %s" % (row + 1, col, text))
                    sheet.write(row + 1, col, text)
                row += 1
            workbook.save(excel_file)
        pass


def main(args):
    print("1. Json to Excel")
    print("2. Excel to Json")
    print("3. Text to Excel")
    # select = input("select: ")
    select = "3"
    if select == "1":
        print("Json to Excel")
        json_file = input("Set Json File: ")
        excel_file = input("Set Excel File: ")
        # json_file = "main.json"
        # excel_file = "out.xls"
        ExcelTools.Json2Excel(json_file, excel_file, sheet_name="sheet1")
    elif select == "2":
        print("Excel to Json")
        excel_file = input("Set Excel File: ")
        input_sheet = input("Set Sheet Index Or Name: ")
        json_file = input("Set Json File(can be null): ")
        print(json_file)
        # json_file = "main.json"
        # excel_file = "out.xls"
        if json_file is not "":
            json_string = ExcelTools.Excel2Json(excel_file, input_sheet, json_file=json_file)
        else:
            json_string = ExcelTools.Excel2Json(excel_file, input_sheet)
        print(json_string)
    elif select == "3":
        print("Text to Excel")
        text_file = "../out/stu.txt"
        excel_file = "../out/out.xls"

        ExcelTools.Text2Excel(text_file, excel_file, "sheet1")
        pass
    else:
        print("Error selection. Exited program.")
        exit(1)


if __name__ == '__main__':
    main(argv)
