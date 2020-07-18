
# xltpl
A python module to generate xls/x files from a xls/x template.


## How it works

xltpl.writer uses xlrd to read xls files, and uses xlwt to write xls files.  
xltpl.writerx uses openpyxl to read and write xlsx files.  
When xltpl reads a xls/x file, it creates a tree for each worksheet.  
Then, it translates the tree to a jinja2 template with custom tags.  
When the template is rendered, jinja2 extensions of cumtom tags call corresponding tree nodes to write the xls/x file.  


xltpl uses jinja2 as its template engine, follows the [syntax of jinja2 template](https://jinja.palletsprojects.com/).  

Each worksheet is translated to a jinja2 template with custom tags.  

```jinja2

{% row 'A1:F4, 0' %}
{% cell '0,0' %}{% endcell %}
{% cell '0,1' %}{% endcell %}
{% cell '0,2' %}{% endcell %}
{% cell '0,3' %}{% endcell %}
{% cell '0,4' %}{% endcell %}
{% cell '0,5' %}{% endcell %}
{% for row in rows %}
{% row 'A2:F2, 3' %}
{% cell '1,0' %}{% endcell %}
{% cell '1,1' %}{% endcell %}
{% cell '1,2' %}{% endcell %}
{% cell '1,3' %}{% endcell %}
{% cell '1,4' %}{% endcell %}
{% cell '1,5' %}{% endcell %}
{% row 'A3:B3, 4' %}
{% cell '2,0' %}{% sec '2,0,0' %}{{loop.index}} {% endsec %}{% endcell %}
{% cell '2,1' %}{% sec '2,1,0' %}{{row[0]}}{% endsec %}{% endcell %}
{% for item in row[1] %}
{% row 'C3, 6' %}
{% cell '2,2' %}{% sec '2,2,0' %}{{item[0]}}{% endsec %}{% endcell %}
{% for mac in item[1] %}
{% row 'D3:E3, 6' %}
{% cell '2,3' %}{% sec '2,3,0' %}{{ mac }}{% endsec %}{% endcell %}
{% cell '2,4' %}{% endcell %}
{% endfor %}
{% endfor %}
{% row 'F3, 4' %}
{% cell '2,5' %}{% endcell %}
{% row 'A4:F4, 3' %}
{% cell '3,0' %}{% endcell %}
{% cell '3,1' %}{% endcell %}
{% cell '3,2' %}{% endcell %}
{% cell '3,3' %}{% endcell %}
{% cell '3,4' %}{% endcell %}
{% cell '3,5' %}{% endcell %}
{% endfor %}
{% row 'A1:F4, 0' %}

```

xltpl added 4 custom tags: row, cell, sec, and xv.  
row, cell, sec are used internally, used for row, cell and rich text.  
xv is used to define a variable.   
When a cell contains only a xv tag, this cell will be set to the type of the object returned from the variable evaluation.  
For example, if a cell contains only {%xv amt %}, and amt is a number, then this cell will be set to Number type, displaying with the style set on the cell.  
If there is another tag, it is equivalent to {{amt}}, will be converted to a string.  



## Installtion

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
  
*   Insert control statements in the notes(comments) of cells, uses beforerow, beforecell or aftercell to seperate them :  
(**v0.4**) You can use 'cell{{A1}}beforerow{% for item in items %}aftercell{% endfor %}'  to specify them.

```jinja2
beforerow{% for item in items %}
```
```jinja2
beforerow{% endfor %}
```

*   (**v0.4**) Use 'range' to specify some regions, using ';;' to seperate them:

```jinja2
range{{D3:E3}}beforerange{% for mac in item[1] %}afterrange{% endfor %};;
range{{C3:E3}}beforerange{% for item in row[1] %}afterrange{% endfor %};;
range{{A2:F4}}beforerange{% for row in rows %}afterrange{% endfor %}
```

*   Run the code
```python
from xltpl.writer import BookWriter
writer = BookWriter('example.xls')
person_info = {'name': u'Hello Wizard'}
items = ['1', '1', '1', '1', '1', '1', '1', '1', ]
person_info['items'] = items
payloads = [person_info]
writer.render_book(payloads)
writer.save('result.xls')
```

See [examples](https://github.com/zhangyu836/python-xlsx-template/tree/master/examples).

## Notes

### Rich text

Openpyxl does not preserve the rich text it read.  
A temporary workaround for rich text is provided in [this repo](https://bitbucket.org/zhangyu836/openpyxl/) ([2.6](https://bitbucket.org/zhangyu836/openpyxl/src/2.6/)).  
For now, xltpl uses this repo to support rich text reading and writing.

### xlrd

xlrd does not extract print settings.   
[This repo](https://github.com/zhangyu836/xlrd) does. 

### xlwt
  
xlwt always sets the default font to 'Arial'.  
Excel measures column width units based on the default font.   
[This repo](https://github.com/zhangyu836/xlwt) does not.  