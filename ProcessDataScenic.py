#!/usr/bin/env python
# -*- coding=utf8 -*-
"""
# File Name    : ProcessDataScenic.py
# Author       : SangYu
# Email        : sangyu.code@gmail.com
# Created Time : 2023年05月22日 星期一 23时05分55秒
# Description  : Process Data Scenic xlsx
"""

import os
import xlwings as xw

SOURCE_DATA_PATH = "source_data"
INPUT_FORMAT = ".xlsx"
OUTPUT_FORMAT = "_process.xlsx"


def process_scenic_data(input_file_path):
    print("process input:{}".format(input_file_path))
    app = xw.App(visible=True, add_book=False)
    work_book = app.books.open(input_file_path)
    sheet0 = work_book.sheets['sheet0']

    # "标题行" --> "开放时间"
    rng = sheet0.range('a1').expand('table')
    nrows = rng.rows.count
    for i in range(nrows):
        value = sheet0[0,i].value
        replace_value = ""
        print(value)
        if value == "标题行":
            replace_value = "开放时间"
        elif value == "字段11":
            replace_value = "景区评级"
        elif value == "字段10":
            replace_value = "景点地址"
        elif value == "字段9":
            replace_value = "主题分类"
        elif value == "字段8":
            replace_value = "本月销量"
        elif value == "字段7":
            replace_value = "价格"
        elif value == "字段3":
            pass
            # sheet0_copy.delete_cols(i)
        elif value == "字段6":
            pass
            # sheet0_copy.delete_cols(i)
        elif value == "字段4":
            replace_value = "标题"
        if replace_value != "":
            sheet0[0, i] = replace_value
        app.quit()

if __name__ == "__main__":
    # demo path
    demo_path = "飞猪湖南scenic.xlsx"
    process_scenic_data(os.path.join(SOURCE_DATA_PATH, demo_path))
    pass
