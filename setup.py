import os
from io import open
from setuptools import setup

CUR_DIR = os.path.abspath(os.path.dirname(__file__))
README = os.path.join(CUR_DIR, "README_EN.md")
with open(README, 'r', encoding='utf-8') as fd:
    long_description = fd.read()

setup(
    name = 'xltpl',
    version = "0.12.1",
    author = 'Zhang Yu',
    author_email = 'zhangyu836@gmail.com',
    url = 'https://github.com/zhangyu836/xltpl',
    packages = ['xltpl'],
	install_requires = ['xlrd >= 1.2.0', 'xlwt >= 1.3.0', 'openpyxl >= 2.6.0', 'jinja2', 'six'],
    description = ( 'A python module to generate xls/x files from a xls/x template' ),
    long_description = long_description,
    long_description_content_type = "text/markdown",
    platforms = ["Any platform "],
    license = 'MIT',
    keywords = ['Excel', 'xls', 'xlsx', 'spreadsheet', 'workbook', 'template']
)