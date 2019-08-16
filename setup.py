from setuptools import setup

setup(
    name = 'xltpl',
    version = "0.1",
    author = 'Zhang Yu',
    author_email = 'zhangyu836@gmail.com',
    packages = ['xltpl'],
	install_requires = ['xlrd', 'xlwt', 'xlutils', 'openpyxl', 'jinja2', 'six'],    
    description = ( 'A python module to generate xls/x files from a xls/x template' ),
    platforms = ["Any platform "],
    license = 'MIT',
    keywords = ['xls', 'xlsx', 'spreadsheet', 'workbook', 'template'],
    python_requires = ">=2.7",
	zip_safe = False
)