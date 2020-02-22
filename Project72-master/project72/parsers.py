import json
import pathlib


def parse_excel(excelfile: pathlib.Path) -> dict:
    """ Parse excel and return a dictionary """


def parse_json(jsonfile: pathlib.Path) -> dict:
    """ Parse a json file and return a dictionary """
    with jsonfile.open("r", encoding="utf-8-sig") as fh:
        data = json.load(fh)

    return data
