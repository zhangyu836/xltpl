# python-xls-xlsx-template
A python module to generate xls/x files from a xls/x template.



## How it works

1 Xltpl reads the xls/x template file and build a tree for each sheet.
2 Xltpl translates the tree to a Jinja2 template with custom tags.
3 When the template is rendered, Jinja2 extensions of cumtom tags call corresponding tree nodes to process rendered content.



## Syntax

You can add Jinja2 template tags: 
  in sheet cells. For example, {{ name }} .
  in cell notes(comments). For example, beforerow{% for item in items %} .
  Beforerow, beforecell, and aftercell are delimiters used to specify where template tag is inserted. 
Technically, any of the Jinja2 syntax can be supported.



## Rich text

Openpyxl does not preserve the rich text it read. 
A temporary workaround for rich text is provided in [this repo](https://bitbucket.org/zhangyu836/openpyxl/) .
For now, Xltpl uses this repo to support rich text reading and writing.



## Dependencies
xlrd, xlwt, xlutils, openpyxl, six, jinja2

