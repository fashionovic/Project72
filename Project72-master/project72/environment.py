from jinja2 import Environment, PackageLoader, select_autoescape

env = Environment(
    loader=PackageLoader(__name__, "templates"),
    autoescape=select_autoescape(["xml"]),
    trim_blocks=True
)

__all__ = ["env"]
