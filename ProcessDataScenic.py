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

import xlrd
import xlwt

SOURCE_DATA_PATH = "source_data"
INPUT_FORMAT = ".xlsx"
OUTPUT_FORMAT = "_process.xlsx"
OUTPUT_ENCODING = "utf-8"

SHEET_FONT_NAME = "宋体"
SHEET_FONT_SIZE = 20 * 11
SHEET_LINE_HEIGHT = (int)(20 * 14.4)

def replace_element_name(head_name_list, old_str, new_str):
    for i in range(len(head_name_list)):
        if head_name_list[i] == old_str:
            head_name_list[i] = new_str


def new_head_line(old):
    """
    input:
       0        1           2        3      4         5        6          7            8          9
    ['链接', '开放时间', '字段3', '标题', '字段6', '价格', '本月销量', '主题分类', '景点地址', '景区评级']
    output:
    ['链接', '景区'，'供应商', '价格', '本月销量', '评价', '评分', '主题分类', '景点地址', '景区评级', '开放时间', '标题']
    """
    return [old[0], "景区", "供应商", old[5], old[6], "评价", "评分", old[7], old[8], old[9], old[1], old[3]]


def new_record_line(old):
    """
    input:
       0        1           2        3      4         5        6          7            8          9
    ['链接', '开放时间', '字段3', '标题', '字段6', '价格', '本月销量', '主题分类', '景点地址', '景区评级']
    output:
    ['链接', '景区'，'供应商', '价格', '本月销量', '评价', '评分', '主题分类', '景点地址', '景区评级', '开放时间', '标题']
    """
    return [old[0], "", "", old[5], old[6], "", "", old[7], old[8], old[9], old[1], old[3]]


def process_scenic_data(input_file_path):
    print("process input:{}".format(input_file_path))
    workbook = xlrd.open_workbook(input_file_path)
    sheet0 = workbook.sheet_by_index(0)
    sheet0_name = sheet0.name

    nrows = sheet0.nrows
    ncols = sheet0.ncols
    print("{} has {} row {} col".format(input_file_path, nrows, ncols))

    head_row_values = sheet0.row_values(0)
    print(head_row_values)

    replace_element_name(head_row_values, "标题行", "开放时间")
    replace_element_name(head_row_values, "字段4", "标题")
    replace_element_name(head_row_values, "字段7", "价格")
    replace_element_name(head_row_values, "字段8", "本月销量")
    replace_element_name(head_row_values, "字段9", "主题分类")
    replace_element_name(head_row_values, "字段10", "景点地址")
    replace_element_name(head_row_values, "字段11", "景区评级")

    head_row_values = new_head_line(head_row_values)
    print(head_row_values)

    # write info to xlsx
    output_file_path = input_file_path.replace(INPUT_FORMAT, OUTPUT_FORMAT)
    print("output_file_path:{}".format(output_file_path))
    write_workbook = xlwt.Workbook(encoding=OUTPUT_ENCODING)
    write_sheet0 = write_workbook.add_sheet(sheet0_name)

    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style.pattern = pattern
    font = xlwt.Font()
    font.name = SHEET_FONT_NAME
    font.height = SHEET_FONT_SIZE
    style.font = font

    style2 = xlwt.XFStyle()
    style2.font = font

    for i in range(nrows):
        write_sheet0.row(i).height = SHEET_LINE_HEIGHT

    for i in range(len(head_row_values)):
        if head_row_values[i] in ["景区","本月销量","标题"]:
            write_sheet0.write(0,i,head_row_values[i],style)
        else:
            write_sheet0.write(0,i,head_row_values[i],style2)


    for i in range(1, nrows):
        line_values = sheet0.row_values(i)
        line_values = new_record_line(line_values)
        # print(line_values)
        for j in range(len(line_values)):
            write_sheet0.write(i,j,line_values[j],style2)

    write_workbook.save(output_file_path)
    print("output_file_path:{} save successful!!!".format(output_file_path))

if __name__ == "__main__":
    # demo path
    demo_path = "飞猪湖南scenic.xlsx"
    process_scenic_data(os.path.join(SOURCE_DATA_PATH, demo_path))
    pass
