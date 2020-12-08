# -*- coding: utf-8 -*-
# @Time    : 2020/12/8 22:52
# @Author  : dashenN72

"""
处理http请求
"""

import requests


def http_request(url, method, param, header=None):
    if method.lower() == 'post':
        r = requests.request('POST', url, headers=header, data=param)
    elif method.lower() == 'get':
        r = requests.request('GET', url, headers=header, params=param)
    else:
        return 0  # 接口请求方法错误
    if r.status_code in (200,):
        return r.text
    else:
        return 1  # 接口状态码错误
