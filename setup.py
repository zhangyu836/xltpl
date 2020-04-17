import os
from io import open
from setuptools import setup

CUR_DIR = os.path.abspath(os.path.dirname(__file__))
README = os.path.join(CUR_DIR, "README.md")
with open(README, 'r', encoding='utf-8') as fd:
    long_description = fd.read()

setup(
    name = 'xltpl',
    version = "0.2",
    author = 'Zhang Yu',
    author_email = 'zhangyu836@gmail.com',
    url = 'https://github.com/zhangyu836/python-xls-xlsx-template',
    packages = ['xltpl'],
	install_requires = ['xlrd', 'xlwt', 'openpyxl', 'jinja2', 'six'],
    description = ( 'A python module to generate xls/x files from a xls/x template' ),
    long_description = long_description,
    long_description_content_type = "text/markdown",
    platforms = ["Any platform "],
    license = 'MIT',
    keywords = ['xls', 'xlsx', 'spreadsheet', 'workbook', 'template']
)