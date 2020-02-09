import json
import pathlib
import typing

from lxml import etree
from jinja2 import Environment, PackageLoader, select_autoescape


__all__: typing.List[str] = ["env", "output_to_xml", "parse_excel"]


env = Environment(
    loader=PackageLoader(__name__, "templates"),
    autoescape=select_autoescape(["xml"]),
    trim_blocks=True
)


def output_to_xml(output: dict) -> str:
    template_name = "kenak.xml.j2"
    template = env.get_template(template_name)
    xml_string = template.render(**output)
    return xml_string


def parse_excel(excelfile: pathlib.Path) -> dict:
    """ Parse excel and return a dictionary """


def parse_json(jsonfile: pathlib.Path) -> dict:
    """ Parse a json file and return a dictionary """
    with jsonfile.open("r") as fh:
        data = json.load(fh)
    return data
