import os
import tkinter
from ttkbootstrap import Style
from tkinter import ttk
import ctypes
import linecache
from tkinter.messagebox import *


if __name__ == '__main__':
    fold_path = os.path.dirname(os.path.abspath(__file__))
    save_fold_path = fold_path + '\\save'
    filepath_wb = save_fold_path + '\\save_wb.txt'
    filepath_sc = save_fold_path + '\\save_sc.txt'

    nW = tkinter.Tk()
    nW.title('设置快捷机器人')  # 窗口标题
    nW.geometry('580x300')  # 窗口尺寸
    style = Style(theme='flatly')
    nW.resizable(False, False)

    # 告诉操作系统使用程序自身的dpi适配
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # 获取屏幕的缩放因子
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
    # 设置程序缩放
    nW.tk.call('tk', 'scaling', ScaleFactor / 75)

    wb_text_1 = tkinter.Label(nW, text="webhook 1")
    wb_text_2 = tkinter.Label(nW, text="webhook 2")
    wb_text_3 = tkinter.Label(nW, text="webhook 3")
    wb_text_1.place(x=5, y=5)
    wb_text_2.place(x=5, y=85)
    wb_text_3.place(x=5, y=165)

    get_wb_1 = tkinter.Entry(nW, width=50)
    get_wb_2 = tkinter.Entry(nW, width=50)
    get_wb_3 = tkinter.Entry(nW, width=50)
    get_wb_1.place(x=95, y=5)
    get_wb_2.place(x=95, y=85)
    get_wb_3.place(x=95, y=165)
    # ————————————————————————————————————————————————————— #
    sc_text_1 = tkinter.Label(nW, text="secret 1")
    sc_text_2 = tkinter.Label(nW, text="secret 2")
    sc_text_3 = tkinter.Label(nW, text="secret 3")
    sc_text_1.place(x=5, y=40)
    sc_text_2.place(x=5, y=120)
    sc_text_3.place(x=5, y=200)

    get_sc_1 = tkinter.Entry(nW, width=50)
    get_sc_2 = tkinter.Entry(nW, width=50)
    get_sc_3 = tkinter.Entry(nW, width=50)
    get_sc_1.place(x=95, y=40)
    get_sc_2.place(x=95, y=120)
    get_sc_3.place(x=95, y=200)
    # ——————————————————————————————————————————————————————— #


    def fill_wb_list():
        get_wb_1.insert(0, linecache.getline(filepath_wb, 1))
        get_wb_2.insert(0, linecache.getline(filepath_wb, 2))
        get_wb_3.insert(0, linecache.getline(filepath_wb, 3))
        get_sc_1.insert(0, linecache.getline(filepath_sc, 1))
        get_sc_2.insert(0, linecache.getline(filepath_sc, 2))
        get_sc_3.insert(0, linecache.getline(filepath_sc, 3))


    def get_wb_list():
        wb_list = [get_wb_1.get(), get_wb_2.get(), get_wb_3.get()]
        return wb_list


    def get_sc_list():
        sc_list = [get_sc_1.get(), get_sc_2.get(), get_sc_3.get()]
        return sc_list


    def write_wb_list():
        wb_list = get_wb_list()
        file = open(filepath_wb, 'w+', encoding='utf-8')
        for line in wb_list:
            file.write(line + '\n')
        file.close()


    def write_sc_list():
        sc_list = get_sc_list()
        file = open(filepath_sc, 'w+', encoding='utf-8')
        for line in sc_list:
            file.write(line + '\n')
        file.close()


    def write_list():
        write_wb_list()
        write_sc_list()
        showinfo('设置成功', '设置成功！')
        print("设置成功")


    fill_wb_list()
    L1 = tkinter.Label(nW, text='')
    L1.place(x=180, y=250)

    set_button = ttk.Button(nW, text="应用", style='success.Outline.TButton', width=30, command=lambda: write_list())
    set_button.place(x=110, y=250)

    nW.bind("<Control-s>", lambda event: write_list())

    nW.mainloop()
