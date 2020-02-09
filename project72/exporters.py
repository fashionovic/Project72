from . import env

def output_to_xml(output: dict) -> str:
    template_name = "kenak.xml.j2"
    template = env.get_template(template_name)
    xml_string = template.render(**output)
    return xml_string


__all__ = ["output_to_xml"]
