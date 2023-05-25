#!/usr/bin/env python
# -*- coding=utf8 -*-
"""
# File Name    : TourismKnowledgeTools.py
# Author       : SangYu
# Email        : sangyu.code@gmail.com
# Created Time : 2023年05月25日 星期四 22时18分14秒
# Description  : Main UI page for tools
"""

import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox

import ParseData5To8

ROOT_TITLE = "TourismKnowledgeTools"
# weightxheight+-x+-y
# +x for screen left,-x for screen right
# +y for screen top,-y for screen bottom
ROOT_GEOMETRY = "1000x600+100+200"


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.button_select_source_file = None
        self.button_quit = None
        self.master = master
        self.grid()

        # create widget here
        self.create_widget()

    def create_widget(self):
        self.button_select_source_file = tk.Button(self, text="选择原始文件并进行处理", command=self.open_source_file)
        self.button_select_source_file.grid(row=0, column=0)
        self.button_quit = tk.Button(self, text="退出", command=root.destroy)
        self.button_quit.grid(row=1, column=0, columnspan=3, sticky=tk.EW)

    def open_source_file(self):
        file = tk.filedialog.askopenfilename(title="选择文件", initialdir="./", filetypes=[("csv文件", ".csv")])
        process_file = ParseData5To8.parse_data(file)
        tk.messagebox.showinfo(title="执行结果", message=process_file+"已生成！！！")


if __name__ == "__main__":
    root = tk.Tk()
    root.title(ROOT_TITLE)
    root.geometry(ROOT_GEOMETRY)
    my_app = Application(master=root)
    my_app.mainloop()
    pass
