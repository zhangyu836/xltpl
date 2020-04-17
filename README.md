
# xltpl
使用 xls/x 文件作为模板来生成 xls/x 文件。  
xltpl 是 [xlstpl](https://github.com/zhangyu836/python-xls-template/) 和 [xlsxtpl](https://github.com/zhangyu836/python-xlsx-template/) 合体。 

## 实现方法

writer 使用 xlrd 来读取 xls 文件，使用 xlwt 来写入 xls 文件。  
wirterx 使用 openpyxl 来读写 xlsx 文件。 
xltpl 读取 xls/x 文件时，会为每个工作表创建一棵树。  
然后，它将树转换为带有自定义 tag 的 jinja2 模板。  
渲染模板时，自定义 tag 所对应的 jinja2 扩展会调用相应的树节点来写入 xls/x 文件。  

### 

xltpl 使用 jinja2 作为模板引擎，遵循 [jinja2 模板的语法](http://docs.jinkan.org/docs/jinja2/templates.html)。  
每个工作表都会被转换为一个带有自定义 tag 的 jinja2 模板。  

```jinja2
...
...
{% row 45 %}
{% cell 46 %}{% endcell %}
{% cell 47 %}{% endcell %}
{% cell 48 %}{{address}}  {%xv v%}{% endcell %}
{% cell 49 %}{% endcell %}
{% cell 50 %}{% endcell %}
{% cell 51 %}{% endcell %}
{% cell 52 %}{% endcell %}
{% cell 53 %}{% endcell %}
{% row 54 %}
{% cell 55 %}{% endcell %}
{% cell 56 %}{% sec 0 %}{{name}}{% endsec %}{% sec 1 %}{{address}}{% endsec %}{% endcell %}
...
...
{% for item in items %}
{% row 64 %}
{% cell 65 %}{% endcell %}
{% cell 66 %}{% endcell %}
{% cell 67 %}{% endcell %}
{% cell 68 %}{% endcell %}
{% cell 69 %}{% endcell %}
{% cell 70 %}{% endcell %}
{% cell 71 %}{% endcell %}
{% cell 72 %}{% endcell %}
{% endfor %}
...
...

```

xltpl 加了 4 个自定义 tag：row、cell、sec 和 xv。  
row、cell、sec 在内部使用，用于处理行、单元格和 rich text。  
xv 用于定义一个变量。一个单元格中只包含 xv 定义的变量时，该单元格会被设置为变量求值所得到的类型。  



## 安装

```shell
pip install xltpl
```

## 使用

*   要使用 xltpl，需要了解 [jinja2 模板的语法](http://docs.jinkan.org/docs/jinja2/templates.html)。  
*   选择一个 xls/x 文件作为模板。  
*   在单元格中插入变量： 
```jinja2
{{name}}
```  
*   在适当的单元格的注释中插入控制语句（使用 beforerow、beforecell 和 aftercell 指定其位置）：
```jinja2
beforerow{% for item in items %}
```
```jinja2
beforerow{% endfor %}
```

*   运行代码
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


[这里](https://github.com/zhangyu836/python-xls-xlsx-template/tree/master/examples)提供了示例。

## 说明

### rich text

openpyxl 读取 rich text 会将它转换为字符串，之后丢弃所读取的 rich text。   
[这里的 openpyxl](https://bitbucket.org/zhangyu836/openpyxl/) ([2.6](https://bitbucket.org/zhangyu836/openpyxl/src/2.6/))会保留 rich text 并支持 rich text 写入。


### xlrd

xlrd 不会读入打印设置。  
如果需要一致的打印设置，可以使用[这里的 xlrd](https://github.com/zhangyu836/xlrd)。 

### xlwt
  
xlwt 总是将默认字体设置为 'Arial'。(Excel 基于默认字体来设置单元格宽度)。   
如果需要一致的单元格宽度，可以使用[这里的 xlwt](https://github.com/zhangyu836/xlwt)。  