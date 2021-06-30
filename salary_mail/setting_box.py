#-*- coding:utf-8 -*-
from datetime import time
import re
import tkinter as tk
import base64
import tkinter.messagebox
from email.header import Header #处理邮件主题
from email.mime.text import MIMEText # 处理邮件内容
from email.utils import parseaddr, formataddr #用于构造特定格式的收发邮件地址
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import  base64
import smtplib #用于发送邮件


class Sender_mail(tk.Toplevel):
    '''邮件发送状态'''
    def __init__(self, parent, sender_detail):
        super(Sender_mail, self).__init__()
        self.title('邮件发送状态')
        cx, cy = parent.get_center()
        self.geometry('300x240+{}+{}'.format(cx - 150, cy - 60))
        self.attributes("-topmost", 1)  # 保持在前
        self.resizable(width=False, height=False)  # 禁制拉伸大小
        self.parent = parent
        self.db = parent.db
        self.sender_detail = sender_detail
        self.setupUI()

    def setupUI(self):
        '''搭建界面'''
        self.thread_count = tk.StringVar()


        # row1 = tk.Frame(self)
        # row1.pack(fill='x', padx=1, pady=5)
        # tk.Label(row1, text="发送线程数：", width=15).pack(side=tk.LEFT)
        # tk.Entry(row1, textvariable=self.thread_count, width=25).pack(side=tk.LEFT)
        
        self.row4 = tk.Text(self,height=15)
        self.row4.pack(fill='x')

        self.row3 = tk.Frame(self)
        self.row3.pack(fill='x')
        tk.Button(self.row3, text='Cancel', command=self.cancel, width=8).pack()
        self.send_run()

    def cancel(self):
        res = tk.messagebox.askyesno(title='是否退出？', message="请检查是否发送完成,如未发送完成,取消可能出现预期外的效果.", parent=self)
        if res:
            self.destroy()


    def check_data(self,datas):
        file_path = datas['send_file_name']
        title = datas['send_title']
        sendtime = datas['send_date']
        from_email = datas['send_from']
        from_name = datas['send_name']
        qm = datas['send_qm']
        main_text = datas['send_main_text']
        send_to = datas['send_to']
        smtp_add = datas['smtp_add']

        if not title:
            return "未检测到邮件标题"
        elif not send_to:
            return "未检测到收件人"
        elif not smtp_add:
            return "未检测到SMTP服务器"
        elif not sendtime:
            return "未检测到发件时间"
        elif not from_email:
            return "未检测到发件箱"
        elif not from_name:
            return "未检测到发件名"
        elif not main_text and not file_path:
            return "未检测到发送内容(附件或正文)"
        
        return 0

    def send_run(self):
        #check
        self.row4.insert("end","校验输入数据...\n")
        self.row4.update()
        check_deail = self.check_data(datas=self.sender_detail)
        if not check_deail:
            self.row4.insert("end","数据校验成功,准备发送..\n")
            self.row4.update()
        else:
            self.row4.insert("end","数据校验失败," + str(check_deail) + ",程序结束...\n")
            self.row4.update()
            return 0

        # "send_file_name": self.salary_file_path.get(),
        # "send_title": self.subject.get(),
        # "send_date": self.send_date.get(),
        # "send_from": self.sender_text.get(),
        # "send_name": self.sender_name_text.get(),
        # "send_qm": self.sign_text.get(),
        # "send_main_text": self.row8.get("1.0",'end-1c')

        from_addr = self.sender_detail["send_from"]
        smtp_server = self.sender_detail["smtp_add"]
        if self.sender_detail["send_main_text"]:
            email_content = self.sender_detail["send_main_text"]
            content = MIMEText(email_content,'html' ,'utf-8')
        else:
            content = MIMEText("",'html' ,'utf-8')
        #附件路径
        msg = MIMEMultipart()
        msg.attach(content)        
        #msg['Received'] = Header("from [127.0.0.1] (unknown [183.129.153.153])by myvps2(Postfix) with ESMTPA id CBBC9101AF2")
        msg['From'] = self.sender_detail["send_from"]
        #msg['From'] = 'zhanghua@ccb.com'
        msg['Subject'] = Header(self.sender_detail["send_title"],'utf-8').encode()
        #邮件带图片
        #file = open('3.png', 'rb')
        #img_data = file.read()
        #file.close()
        #img = MIMEImage(img_data,_subtype='octet-stream')
        #img.add_header('Content-ID', '<image1>')
        #msg.attach(img)
        #带附件
        if self.sender_detail["send_file_name"]:
            docfile = self.sender_detail["send_file_name"]
            docpart = MIMEApplication(open(docfile, 'rb').read())
            docpart.add_header('Content-Disposition', 'attachment', filename = docfile)
            msg.attach(docpart)
        
        
        #print(to_addr.strip())
        msg['To'] = self.sender_detail["send_to"]
        #print(msg['To'])
        
        server = smtplib.SMTP(smtp_server,25)
        #server.login(from_addr,password)
        try:
            server.sendmail(from_addr, self.sender_detail["send_to"], msg.as_string())
            self.row4.insert("end",'发送成功--' + self.sender_detail["send_to"] + "\n")
            self.row4.update()
        except Exception as e:
            self.row4.insert("end","发送失败\n")
            self.row4.update()
            self.row4.insert("end", e)      
            self.row4.update()

            
            #server.sendmail(from_addr, to_addr, msg.as_string())
        #print('发送邮件到'+to_addr)
        server.quit()



        return 1


