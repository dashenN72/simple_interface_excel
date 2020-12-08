# -*- coding: utf-8 -*-
# @Time    : 2020/12/8 22:52
# @Author  : dashenN72
"""
入口方法
"""

from op_excel import ExcelUtil
from op_http import http_request

excelUtil = ExcelUtil()
path = './data/xxx接口测试_参数校验.xlsx'
excelUtil.load_excel(path)  # 加载excel
excelUtil.get_sheet_by_name("参数校验")
rows = excelUtil.get_sheet_rows()  # 获取行数
for row in range(1, rows):
    data = excelUtil.get_row_value(row)  # 获取行数据
    print("[INFO]第[%d]行待处理数据：%s" % (row, str(data)))
    result = http_request(data[3], data[2], data[4])
    print("[INFO]接口返回数据：%s" % str(result))
    if result == 0:
        print("[ERROR]不支持的接口请求方法")
    elif result == 1:
        print("[ERROR]接口状态码错误")
    else:
        if excelUtil.write_data(row=row + 1, col=7, value=result, path=path):
            print("[INFO]写数据到excel成功")
        else:
            print("[ERROR]写数据到excel失败")
