
# xltpl
A python module to generate xls/x files from a xls/x template.

## How it works

When xltpl reads a xls/x file, it creates a tree for each worksheet.  
And, each tree is translated to a jinja2 template with custom tags.  
When the template is rendered, jinja2 extensions of cumtom tags call corresponding tree nodes to write the xls/x file.

## How to install

```shell
pip install xltpl
```

## How to use

*   To use xltpl, you need to be familiar with the [syntax of jinja2 template](https://jinja.palletsprojects.com/).
*   Get a pre-written xls/x file as the template.
*   Insert variables in the cells, such as : 

```jinja2
{{name}}
```
  
*   ~~Insert control statements in the notes(comments) of cells, use beforerow, beforecell or aftercell to seperate them :~~


```jinja2
beforerow{% for item in items %}
```
```jinja2
beforerow{% endfor %}
```

*   Insert control statements in the cells (**v0.9**) :

```jinja2
{%- for row in rows %}
{% set outer_loop = loop %}{% for row in rows %}
Cell
{{outer_loop.index}}{{loop.index}}
{%+ endfor%}{%+ endfor%}
```

*   Run the code
```python
from xltpl.writerx import BookWriter
writer = BookWriter('tpl.xlsx')
person_info = {'name': u'Hello Wizard'}
items = ['1', '1', '1', '1', '1', '1', '1', '1', ]
person_info['items'] = items
payloads = [person_info]
writer.render_book(payloads)
writer.save('result.xlsx')
```

## Supported
* MergedCell   
* Non-string value for a cell (use **{% xv variable %}** to specify a variable) 
* For xlsx  
Image (use **{% img variable %}**)  
DataValidation   
AutoFilter


## Related
* [pydocxtpl](https://github.com/zhangyu836/pydocxtpl)  
A python module to generate docx files from a docx template.
* [django-excel-export](https://github.com/zhangyu836/django-excel-export)  
A Django library for exporting data in xlsx, xls, docx format, utilizing xltpl and pydocxtpl, with admin integration.  
[Demo project](https://github.com/zhangyu836/django-excel-export-demo)   
[Live demo](https://tranquil-tundra-83829.herokuapp.com/) (User name: admin
Password: admin)   

* [xltpl for nodejs](https://github.com/zhangyu836/node-xlsx-template)
* [xltpl for java](https://github.com/zhangyu836/xltpl4java)


## Notes

### xlrd

xlrd does not extract print settings.   
[This repo](https://github.com/zhangyu836/xlrd) does. 

### xlwt
  
xlwt always sets the default font to 'Arial'.  
Excel measures column width units based on the default font.   
[This repo](https://github.com/zhangyu836/xlwt) does not.  
