import pathlib

from project72 import output_to_xml, parse_excel, parse_json


def main():
    output_file = pathlib.Path("output.xml")
    input_file = pathlib.Path("sample_input.json")
    data = parse_json(input_file)
    xml = output_to_xml(data)
    output_file.write_text(xml)


if __name__ == "__main__":
    main()
