# -*- coding: utf-8 -*-
import xlrd
import json
import os

import shutil


def read_template(template="excel2py.json"):
    with open(template, "r") as fd:
        data = json.load(fd)
    return data


def parse_xlsx(filename):
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)
    no_rows = sh.nrows
    parsed_data = {}
    parsed_data["opaque_rows"] = no_rows
    parsed_data["opaque_column0"] = ""
    for column_index in range(1, 17):
        key = f"opaque_column{column_index}"
        try:
            values = sh.col_values(column_index - 1)
            values = ",".join(str(v) for v in values)
        except IndexError:
            values = "," * no_rows
        parsed_data[key] = values
    return parsed_data


def generate_json(excel_file, output_file):
    data = read_template()
    parsed_data = parse_xlsx(excel_file)
    data.update(**parsed_data)
    with open(output_file, "w") as fd:
        json.dump(data, fd, indent=2)


generate_json("calc.xlsx", "out.json")
