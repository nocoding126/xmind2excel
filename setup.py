"""
!/usr/bin/env python
# -*- coding: utf-8 -*-
@Time    : 2023/5/26 17:39
@Author  : 派大星
@Site    : 
@File    : setup.py
@Software: PyCharm
@desc:
"""
import os
from xmind2excel.xmind_to_excel import get_xmind_data, xmind_to_list, write_excel


def find_xmind():
    """在当前项目中查找XMind文件"""
    files_path = []
    xmind_path = []
    for _dir, dirs, files in os.walk("."):

        for file in files:
            file_path = os.path.join(_dir, file)
            files_path.append(file_path)

    for _path in files_path:
        if _path.endswith(".xmind"):
            xmind_path.append(_path)
    return xmind_path


def xmind_to_excel():
    """把xmind文件数据写入Excel"""
    xmind_name = input("请输入xmind文件名")
    xmind_path = find_xmind()
    for xmind in xmind_path:
        if xmind_name in xmind:
            data_list = get_xmind_data(xmind)
            write_excel(*xmind_to_list(data_list))
        else:
            print("xmind文件不存在！")


if __name__ == '__main__':
    xmind_to_excel()
