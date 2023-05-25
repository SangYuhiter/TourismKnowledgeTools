#!/usr/bin/env python
# -*- coding=utf8 -*-
"""
# File Name    : ParseData5-8.py
# Author       : SangYu
# Email        : sangyu.code@gmail.com
# Created Time : 2023年05月22日 星期一 21时56分06秒
# Description  : Parse Data in guide from section 5 to section 8
"""

import os
import xlwt

# SOURCE_DATA_PATH = "./source_data"
SOURCE_DATA_PATH = "."
INPUT_FORMAT = ".csv"
OUTPUT_FORMAT = "_process.xlsx"
INPUT_ENCODING = "gbk"
OUTPUT_ENCODING = "utf-8"

SCENIC_KEY_WORD = "scenic"
TRAVELDETAIL_KEY_WORD = "traveldetail"

SCENIC_SHEET_NAME = "scenic"
TRAVELDETAIL_SHEET_NAME = "td"

SHEET_FONT_NAME = "宋体"
SHEET_FONT_SIZE = 20 * 11
SHEET_LINE_HEIGHT = (int)(20 * 14.4)


def parse_data(input_file_path):
    print("process input:{}".format(input_file_path))
    # use set for remove duplicate record
    all_lines = set()
    with open(input_file_path, "r", encoding=INPUT_ENCODING) as fr:
        # remove line 0
        for line in fr.readlines()[1:]:
            for record in line.split(";"):
                all_lines.add(record.strip())

    # sort records
    all_lines = sorted(all_lines)

    # split records
    scenic_records = []
    travel_detail_records = []
    for line in all_lines:
        if SCENIC_KEY_WORD in line:
            scenic_records.append(line)
        elif TRAVELDETAIL_KEY_WORD in line:
            travel_detail_records.append(line)
        else:
            print("can not cover line:{}", line)

    for line in all_lines:
        # print(line)
        pass
    print("all record size:{},scenic record size:{},traveldetail record size:{}".format(len(all_lines),
                                                                                        len(scenic_records),
                                                                                        len(travel_detail_records)))

    # write info to xlsx
    output_file_path = input_file_path.replace(INPUT_FORMAT, OUTPUT_FORMAT)
    print("output_file_path:{}".format(output_file_path))
    work_book = xlwt.Workbook(encoding=OUTPUT_ENCODING)
    travel_detail_sheet = work_book.add_sheet(TRAVELDETAIL_SHEET_NAME)
    scenic_sheet = work_book.add_sheet(SCENIC_SHEET_NAME)
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = SHEET_FONT_NAME
    font.height = SHEET_FONT_SIZE
    style.font = font
    # bypass head line,just save data
    for i in range(len(travel_detail_records)):
        travel_detail_sheet.row(i).height = SHEET_LINE_HEIGHT
        travel_detail_sheet.write(i, 0, travel_detail_records[i], style)
    for i in range(len(scenic_records)):
        scenic_sheet.row(i).height = SHEET_LINE_HEIGHT
        scenic_sheet.write(i, 0, scenic_records[i], style)
    work_book.save(output_file_path)
    print("output_file_path:{} save successful!!!".format(output_file_path))
    return output_file_path


if __name__ == "__main__":
    # demo path
    for file in os.listdir(SOURCE_DATA_PATH):
        if INPUT_FORMAT in file:
            parse_data(os.path.join(SOURCE_DATA_PATH, file))
    input("请按任意键退出......")
    pass
