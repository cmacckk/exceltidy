# Excel Tidy

[![PyPI version](https://badge.fury.io/py/exceltidy.svg)](https://badge.fury.io/py/exceltidy)

在工作中遇到的一些需要处理的情况,编写为库,方便后续操作

`openpyxl`前缀的函数为基于`openpyxl`编写
`xlwings`同上

注意`xlwings`不支持`linux`,在`linux`上运行可能报错

使用方法
`git clone https://github.com/cmacckk/exceltidy.git`

安装依赖
`pip install -r requirements.txt`

`openpyxl`获取`list[list, list, ...]` `openpyxl.workbook.Workbook` `openpyxl.worksheet.worksheet.Worksheet`
```python
from exceltidy.exceltidy import *

worksheet_datas, workbook, worksheet = openpyxl_get_datas_workbook_worksheet(filename="./test.xlsx")

# 通过工作表获取数据
datas = openpyxl_get_datas(worksheet)

# 替换工作表数据
openpyxl_replace_worksheet_data(origin_worksheet, to_be_replaced_worksheet)

# 保存及关闭工作薄
workbook.save('./result.xlsx')
workbook.close()
```

`xlwings`获取`App` `xw.Book` `xw.Sheet`
`sheet_name`为`None`时读取返回第一个工作表
```python
from exceltidy.exceltidy import *

app, workbook, worksheet = xlwings_get_app_workbook_worksheet('./test.xlsx')

# 
_, _, to_be_replaced_worksheet = xlwings_get_app_workbook_worksheet('./test2.xlsx', sheet_name='sort')

# 替换工作表数据
xlwings_replace_worksheet_data(origin_worksheet, to_be_replaced_worksheet)

# 排序
xlwings_sort(worksheet, "A4:L11", "L4:L11", 'descending')

# 保存及关闭工作薄、App
workbook.save()
workbook.close()
app.quit()
```

`openpyxl_parse_multi_title_to_single`及`openpyxl_parser_merged_cell`使用方法详见代码函数