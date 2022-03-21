
# xltpl  
使用 xls/x 文件作为模板来生成 xls/x 文件。 [English](README_EN.md)     

## 实现方法

xls/x 文件的每个工作表会被转换为一棵树。  
树会被转换为带有自定义 tag 的 jinja2 模板。  
渲染模板时，自定义 tag 所对应的 jinja2 扩展调用相应的树节点来写入 xls/x 文件。  

 
## 安装

```shell
pip install xltpl
```

## 使用

*   要使用 xltpl，需要了解 [jinja2 模板的语法](http://docs.jinkan.org/docs/jinja2/templates.html) 。  
*   选择一个 xls/x 文件作为模板。  
*   在单元格中插入变量： 
```jinja2
{{name}}
```  
*   ~~在单元格的批注中插入控制语句（使用 beforerow、beforecell 和 aftercell 指定其位置）：~~  


```jinja2
beforerow{% for item in items %}
```
```jinja2
beforerow{% endfor %}
```

*   或在单元格中插入控制语句(**v0.9**)：

```jinja2
{%- for row in rows %}
{% set outer_loop = loop %}{% for row in rows %}
Cell
{{outer_loop.index}}{{loop.index}}
{%+ endfor%}{%+ endfor%}
```


*   运行代码
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

## 支持的特性
* 合并单元格 (MergedCell)   
* 单元格非字符串值 (使用 **{% xv variable %}** 来表示变量)   
* 对于 xlsx     
图片 (使用 **{% img variable %}**)     
数据有效性(DataValidation)     
筛选 (AutoFilter)   


## 相关
* [pydocxtpl](https://github.com/zhangyu836/pydocxtpl)  
使用 docx 文件作为模板来生成 docx 文件。  
其实现方法与 xltpl 类似。
* [django-excel-export](https://github.com/zhangyu836/django-excel-export)  
利用 xltpl 和 pydocxtpl 在 Django admin 后台以xls/x 和 docx 格式导出数据。  
[演示项目](https://github.com/zhangyu836/django-excel-export-demo)  
[在线演示](https://tranquil-tundra-83829.herokuapp.com/) (用户名: admin
密码: admin)   
* [nodejs 版本的 xltpl](https://github.com/zhangyu836/node-xlsx-template)   
CodeSandbox examples: 
[browser](https://codesandbox.io/s/xlsx-export-with-exceljs-and-xltpl-58j9g6)
[node](https://codesandbox.io/s/exceljs-template-with-xltpl-4w58xo)   
* [xltpl for java](https://github.com/zhangyu836/xltpl4java)


## 说明

### xlrd

xlrd 不会读入打印设置。  
如果需要一致的打印设置，可以使用[这里的 xlrd](https://github.com/zhangyu836/xlrd) 。 

### xlwt
  
xlwt 总是将默认字体设置为 'Arial'。  
Excel 基于默认字体来设置单元格宽度。   
如果需要一致的单元格宽度，可以使用[这里的 xlwt](https://github.com/zhangyu836/xlwt) 。  
