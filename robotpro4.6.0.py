import base64
import ctypes  # 缩放适配
import hashlib
import hmac
import json
import pathlib
import time
import tkinter.messagebox  # 消息事件库
import urllib.parse
import webbrowser
import os
from tkinter import ttk
import requests
from ttkbootstrap import Style
import linecache

if __name__ == "__main__":

    nW = tkinter.Tk()
    nW.title('钉钉机器人pro  4.6.0')  # 窗口标题
    nW.geometry('580x460')  # 窗口尺寸
    nW.resizable(False, False)
    style = Style(theme='flatly')

    # 告诉操作系统使用程序自身的dpi适配
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # 获取屏幕的缩放因子
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
    # 设置程序缩放
    nW.tk.call('tk', 'scaling', ScaleFactor / 75)

    fold_path = os.path.dirname(os.path.abspath(__file__))
    save_fold_path = fold_path + '\\save'
    filepath_wb = save_fold_path + '\\save_wb.txt'
    filepath_sc = save_fold_path + '\\save_sc.txt'
    list_num = tkinter.IntVar()
    list_num.set(0)


    def get_msgtype():
        message_type = msgtype1.get()
        if message_type == 1:
            msgtype = "text"
        else:
            msgtype = "markdown"
        return msgtype


    def open_url():
        webbrowser.open("https://open.dingtalk.com/document/orgapp/custom-robots-send-group-messages", new=0)


    msgtype1 = tkinter.IntVar()
    msgtype1.set(1)

    word1 = tkinter.Label(nW, text='webhook')
    word2 = tkinter.Label(nW, text='加签密匙')
    word3 = tkinter.Label(nW, text='内容')
    wordIntro = tkinter.Label(nW, text='made by AXIOS with Python @4.6.0', font=('Arial', 7, 'italic'))
    word1.place(x=480, y=80)
    word2.place(x=480, y=120)
    word3.place(x=480, y=160)
    wordIntro.place(x=150, y=430)

    R1 = tkinter.Radiobutton(nW, text="纯文本模式", variable=msgtype1, value=1, command=get_msgtype())
    R2 = tkinter.Radiobutton(nW, text="MarkDown模式(beta)", variable=msgtype1, value=2, command=get_msgtype())
    R2Button = ttk.Button(nW, text="查看指南", width=10, style='success-link', command=lambda: open_url())
    R1.place(x=15, y=10)
    R2.place(x=140, y=10)
    R2Button.place(x=325, y=9)

    t1 = tkinter.Entry(nW, width=50)
    t2 = tkinter.Entry(nW, width=50)
    t3 = tkinter.Text(nW, width=49, height=10)
    t1.place(x=15, y=80)
    t2.place(x=15, y=120)
    t3.place(x=15, y=160)

    p1 = tkinter.ttk.Progressbar(nW, orient=tkinter.HORIZONTAL, length=580, mode='determinate')
    p1.place(x=0, y=450)
    p1['maximum'] = 100
    p1['value'] = 0


    def bar_move():
        for i in range(20):
            # 每次更新加1
            p1.step(1)
            # 更新画面
            nW.update()
            time.sleep(0.001)


    def write_entry():
        num_list = list_num.get()
        if num_list == 0:
            t1.delete(0, 'end')
            t1.insert(0, linecache.getline(filepath_wb, 1).replace('\n', '').replace('\r', ''))
            t2.delete(0, 'end')
            t2.insert(0, linecache.getline(filepath_sc, 1).replace('\n', '').replace('\r', ''))
        elif num_list == 1:
            t1.delete(0, 'end')
            t1.insert(0, linecache.getline(filepath_wb, 2).replace('\n', '').replace('\r', ''))
            t2.delete(0, 'end')
            t2.insert(0, linecache.getline(filepath_sc, 2).replace('\n', '').replace('\r', ''))
        elif num_list == 2:
            t1.delete(0, 'end')
            t1.insert(0, linecache.getline(filepath_wb, 3).replace('\n', '').replace('\r', ''))
            t2.delete(0, 'end')
            t2.insert(0, linecache.getline(filepath_sc, 3).replace('\n', '').replace('\r', ''))
        else:
            print("write_entry error")
        return "1"


    def open_rs():
        os.system('python ' + fold_path + '\\RobotSave.py')


    def get_home_dir():
        homedir = str(pathlib.Path.home())
        return homedir


    def get_url():
        url = t1.get()
        return url


    def get_secret():
        secret = t2.get()
        return secret


    def get_timestamp_sign():
        timestamp = str(round(time.time() * 1000))
        secret = get_secret()
        secret_enc = secret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, secret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))

        bar_move()
        print("timestamp: ", timestamp)
        print("sign:", sign)
        return timestamp, sign


    def get_signed_url():
        url = get_url()
        timestamp, sign = get_timestamp_sign()
        webhook = url + "&timestamp=" + timestamp + "&sign=" + sign
        bar_move()
        return webhook


    def get_webhook(mode):
        url = get_url()
        if mode == 0:  # only 敏感字
            webhook = url
        elif mode == 1 or mode == 2:  # 敏感字和加签 或 # 敏感字+加签+ip
            webhook = get_signed_url()
        else:
            webhook = ""
            print("error! mode:   ", mode, "  webhook :  ", webhook)
        bar_move()
        return webhook


    def get_message(content, is_send_all):
        # 和类型相对应，具体可以看文档 ：https://ding-doc.dingtalk.com/doc#/serverapi2/qf2nxq 
        # 可以设置某个人的手机号，指定对象发送
        global message
        msgtype = get_msgtype()
        if msgtype == "text":
            message = {
                "msgtype": "text",  # 有text, "markdown"、link、整体跳转ActionCard 、独立跳转ActionCard、FeedCard类型等
                "text": {
                    "content": content  # 消息内容
                },
                "at": {
                    "atMobiles": [
                        "1862*8*****6",
                    ],
                    "isAtAll": is_send_all  # 是否是发送群中全体成员
                }
            }
        elif msgtype == "markdown":
            message = {
                "msgtype": "markdown",  # 有text, "markdown"、link、整体跳转ActionCard 、独立跳转ActionCard、FeedCard类型等
                "markdown": {
                    "title": "123",
                    "text": content  # 消息内容
                },
                "at": {
                    "isAtAll": is_send_all  # 是否是发送群中全体成员
                }
            }
        bar_move()
        print(message)
        print(msgtype)
        return message


    def send_ding_message(content, is_send_all):
        # 请求的URL，WebHook地址
        webhook = get_webhook(1)  # 主要模式有 0 ： 敏感字 1：# 敏感字 +加签 3：敏感字+加签+IP
        print("webhook: ", webhook)
        # 构建请求头部
        header = {
            "Content-Type": "application/json",
            "Charset": "UTF-8"
        }
        # 构建请求数据
        message = get_message(content, is_send_all)
        # 对请求的数据进行json封装
        message_json = json.dumps(message)
        # 发送请求
        info = requests.post(url=webhook, data=message_json, headers=header)
        bar_move()
        # 打印返回的结果
        print(info.text)


    def delete():
        t3.delete(0.0, "end")


    def send_message():
        content = t3.get(0.0, 'end-1c')
        is_send_all = "False"
        send_ding_message(content, is_send_all)
        delete()
        p1['value'] = 0


    R1_write_robot = tkinter.Radiobutton(nW, text="快捷-1", variable=list_num, value=0, command=lambda: write_entry())
    R2_write_robot = tkinter.Radiobutton(nW, text="快捷-2", variable=list_num, value=1, command=lambda: write_entry())
    R3_write_robot = tkinter.Radiobutton(nW, text="快捷-3", variable=list_num, value=2, command=lambda: write_entry())
    R1_write_robot.place(x=15, y=45)
    R2_write_robot.place(x=105, y=45)
    R3_write_robot.place(x=195, y=45)

    b1 = ttk.Button(nW, text="发送消息", style='success.Outline.TButton', width=30,command=lambda: send_message())
    b2 = ttk.Button(nW, text="设置快捷机器人", width=18, style='success-link', command=lambda: open_rs())
    b2.place(x=285, y=45)
    b1.place(x=110, y=390)

    nW.bind("<Control-Return>", lambda event: send_message())

    nW.mainloop()
