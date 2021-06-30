from os import name
import threading
import datetime
import base64

import tkinter as tk
from email.mime.text import MIMEText
from email.utils import formataddr

from tkinter import ttk
from .db_instance import set_db, SalaryEmail
from .setting_box import Sender_mail
import tkinter.messagebox
import tkinter.filedialog
import re
from smtplib import SMTP_SSL, SMTP
from decimal import (Decimal)


class MainWin(tk.Tk):

    def __init__(self):
        super(MainWin, self).__init__()
        self.title('mail sender  -- by s-hak')
        x = (self.winfo_screenwidth() // 2) - 300
        y = (self.winfo_screenheight() // 2) - 300
        self.geometry('600x600+{}+{}'.format(x, y))
        self.resizable(width=False, height=False)  # 禁制拉伸大小
        self.label_width = 55  # 标签长度

        self.db = set_db()
        self.subject = tk.StringVar() # 邮件标题
        self.salary_file_path = tk.StringVar()
        self.send_date = tk.StringVar() # 发件时间
        self.sender_text = tk.StringVar()
        self.sender_name_text = tk.StringVar()
        self.sign_text = tk.StringVar()
        self.set_text = tk.StringVar()
        self.send_to = tk.StringVar()
        self.smtp_add = tk.StringVar()
        self.__password = tk.StringVar()


        self.excel_file = None
        self.setupUI()
        self.set_default_info()

    def setupUI(self):
        '''主界面'''

        # 设置菜单栏
        menubar = self.set_menubar()
        self.config(menu=menubar)
        self.show_base_info()

    # def show_info_box(self):
    #     info_box = InfoWin(parent=self)
    #     self.wait_window(info_box)

    def set_menubar(self):
        '''菜单栏'''
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="文件", menu=filemenu)

        filemenu.add_command(label='退出', command=self.quit)
        settingmenu = tk.Menu(menubar, tearoff=False)
        menubar.add_cascade(label="设置", menu=settingmenu)
        # settingmenu.add_command(label="账号/密码", command=self.show_account_box)
        # settingmenu.add_command(label="SMTP域名/端口", command=self.show_smtp_port_box)
        # settingmenu.add_command(label="邮件信息设置", command=self.show_info_box)
        # settingmenu.add_command(label="系统设置", command=self.show_sys_setting_box)
        settingmenu.add_command(label="预留", command="")
        return menubar


    def show_base_info(self):
        '''显示基本信息'''

        row1 = tk.Frame(self)
        row1.pack(fill='x', padx=1, pady=5)
        tk.Label(row1, text="发件邮箱：", width=15).pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.sender_text, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)
        row2 = tk.Frame(self)
        row2.pack(fill='x', padx=1, pady=5)
        tk.Label(row2, text="发件人：", width=15).pack(side=tk.LEFT)
        tk.Entry(row2, textvariable=self.sender_name_text, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Button(row2, text='生成', command=self.set_from_name, width=5).pack(side=tk.LEFT)


        row9 = tk.Frame(self)
        row9.pack(fill='x', padx=1, pady=5)
        tk.Label(row9, text="收件邮箱：", width=15).pack(side=tk.LEFT)
        tk.Entry(row9, textvariable=self.send_to, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)

        row10 = tk.Frame(self)
        row10.pack(fill='x', padx=1, pady=5)
        tk.Label(row10, text="SMTP服务器", width=15).pack(side=tk.LEFT)
        tk.Entry(row10, textvariable=self.smtp_add, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Button(row10, text='识别', command=self.cat_mx, width=5).pack(side=tk.LEFT)

        row3 = tk.Frame(self)
        row3.pack(fill='x', padx=1, pady=5)
        tk.Label(row3, text="邮件标题：", width=15).pack(side=tk.LEFT)
        tk.Entry(row3, textvariable=self.subject, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)

        row5 = tk.Frame(self)
        row5.pack(fill='x', padx=1, pady=5)
        tk.Label(row5, text="邮件签名/落款：", width=15).pack(side=tk.LEFT)
        tk.Entry(row5, textvariable=self.sign_text, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)

        row6 = tk.Frame(self)
        row6.pack(fill='x', padx=1, pady=5)
        tk.Label(row6, text="邮件日期：", width=15).pack(side=tk.LEFT)
        tk.Entry(row6, textvariable=self.send_date, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)

        row4 = tk.Frame(self)
        row4.pack(fill='x', padx=1, pady=5)
        tk.Label(row4, text="附件：", width=15).pack(side=tk.LEFT)
        tk.Entry(row4, textvariable=self.salary_file_path, width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)
        tk.Button(row4, text='Open', command=self.get_salary_file_path, width=5).pack(side=tk.LEFT)
        
        row7 = tk.Frame(self)
        row7.pack(fill='x', padx=1, pady=5)
        tk.Label(row7, text="正文：", width=15).pack(side=tk.LEFT)


        self.row8 = tk.Text(self,height=18)
        self.row8.pack(fill='x', padx=1, pady=5)
        
        
        # tk.Label(row8, text="正文：", width=15).pack(side=tk.LEFT)
        # tk.Entry(row8, textvariable="", width=self.label_width, justify=tk.CENTER).pack(side=tk.LEFT)
        

        tk.Button(self, command=self.send_email, text='发送', width=20).pack(padx=1, pady=5)


    def set_from_name(self):
        name_all = self.sender_text.get()
        out = str(re.findall(".*@",name_all)[0]).replace("@","")
        self.sender_name_text.set(out)



    def cat_mx(self):
        try:
            import dns.resolver
            catdomain = self.send_to.get()
            domain = str(re.findall("@.*",catdomain)[0]).replace("@","")
            mx= dns.resolver.query(domain,'MX')
            for i in mx.response.answer[0]:
                a = i.exchange.labels
                allstr = ""
                for i in a:
                    allstr = allstr + str(i)
                    "b'163mx02'b'mxmail'b'netease'b'com'b''"
                outs = str(allstr).replace("'b''","").replace("'b'",".").replace("b'","").replace("'","")
                self.smtp_add.set(outs)
                return 1
        except:
            pass


    def get_salary_file_path(self):
        '''获取文件路径'''
        path = tk.filedialog.askopenfilename(title='选择文件', filetypes=[("upload File", "")])
        self.salary_file_path.set(path)

    def set_default_info(self):
        '''设置默认初始值'''
        self.send_date.set(datetime.datetime.now().strftime("%Y-%m-%d"))
        



    def send_email(self):
        '''发送邮件'''

        sender_detail = {
            "send_file_name": self.salary_file_path.get(),
            "send_title": self.subject.get(),
            "send_date": self.send_date.get(),
            "send_from": self.sender_text.get(),
            "send_name": self.sender_name_text.get(),
            "send_qm": self.sign_text.get(),
            "send_main_text": self.row8.get("1.0",'end-1c'),
            "send_to": self.send_to.get(),
            "smtp_add": self.smtp_add.get()
        }

        sys_setting_box = Sender_mail(parent=self,sender_detail=sender_detail)
        self.wait_window(sys_setting_box)
        return 1


    def get_center(self):
        px = self.winfo_x()
        py = self.winfo_y()
        pw = self.winfo_width()
        ph = self.winfo_height()

        return (int(px + pw/2), int(py + ph/2))



if __name__ == '__main__':
    win = MainWin()
    win.mainloop()
