#-*- coding:utf-8 -*-
from email.header import Header #处理邮件主题
from email.mime.text import MIMEText # 处理邮件内容
from email.utils import parseaddr, formataddr #用于构造特定格式的收发邮件地址
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import  base64
import smtplib #用于发送邮件
#def _format_addr(s):
#	name, addr = parseaddr(s)
#	return formataddr((Header(name, 'utf-8').encode(), addr))



with open('email_list.txt','r+',encoding='utf-8') as f:
	to_addrs = f.readlines()
for to_addr in to_addrs:
	from_addr = 'zh_l@jslhbank.com'
	#password = 'XXXXX'
	#smtp_server = 'mail.jshbank.com'
	smtp_server = '163mx01.mxmail.netease.com'
	#url='http://27.102.130.158/longxingsihai.exe'
	email_content= "<style class='fox_global_style'>div.fox_html_content { line-height: 1.5; }p { margin-top: 0px; margin-bottom: 0px; }div.fox_html_content { font-size: 14px; font-family: 'Microsoft YaHei UI'; color: rgb(0, 0, 0); line-height: 1.5; }</style><div><span id='_FoxCURSOR'></span>周老师：</div><div><span microsoft='' yahei='' ui';='' font-size:='' 14px;='' color:='' rgb(0,='' 0,='' 0);='' background-color:='' rgba(0,='' font-weight:='' normal;='' font-style:='' normal;text-decoration:='' none;'=''>&nbsp; &nbsp; 您好！测试方案已经发送至您的邮箱，请查收。</span></div><div><span microsoft='' yahei='' ui';='' font-size:='' 14px;='' color:='' rgb(0,='' 0,='' 0);='' background-color:='' rgba(0,='' font-weight:='' normal;='' font-style:='' normal;text-decoration:='' none;'=''><span microsoft='' yahei='' ui';='' font-size:='' 14px;='' color:='' rgb(0,='' 0,='' 0);='' background-color:='' rgba(0,='' font-weight:='' normal;='' font-style:='' normal;text-decoration:='' none;'=''>&nbsp; &nbsp; 有问题您在联系我</span></span></div><div><br></div><hr id='FMSigSeperator' style='width: 210px; height: 1px;' color='#b5c4df' size='1' align='left'><div><span id='_FoxFROMNAME'><div style='MARGIN: 10px; FONT-FAMILY: verdana; FONT-SIZE: 10pt'><br></div><div style='MARGIN: 10px; FONT-FAMILY: verdana; FONT-SIZE: 10pt'><div style='font-family: Verdana; font-size: 14px;'><div style='width: 480px; position: relative; border-top: 2px dashed rgb(221, 221, 221);'><p style='margin-right: 0px; margin-left: 0px; padding-top: 15px; line-height: 22px;'><span style='font-weight: bolder; font-size: 22px; margin-right: 10px;'>王智超</span>安全服务工程师</p><p style='color: rgb(24, 161, 238); font-size: 18px; letter-spacing: 2px;'>北京中睿天下信息技术有限公司</p><p style='margin-right: 0px; margin-left: 0px; color: rgb(51, 51, 51); line-height: 22px;'>电话：16601110024</p><p style='margin-right: 0px; margin-left: 0px; color: rgb(51, 51, 51); line-height: 22px;'>邮箱：<a href='mailto:chentian@zorelworld.com' target='_blank' style='outline: none; color: rgb(44, 74, 119);'>wangzhichao@zo<wbr>relworld.com</a></p><p style='margin-right: 0px; margin-left: 0px; color: rgb(51, 51, 51); line-height: 22px;'>网址：<a href='http://www.zorelworld.com/' target='_blank' style='outline: none; color: rgb(44, 74, 119);'>http://www.z<wbr>orelworld.co<wbr>m/</a></p><p style='margin-right: 0px; margin-left: 0px; color: rgb(51, 51, 51); line-height: 22px;'>地址：北京市海淀区农大南路1号院硅谷亮城2B座212-215</p><p style='color: rgb(51, 51, 51); height: 30px; line-height: 30px; width: 480px; border-top: 2px dashed rgb(221, 221, 221); border-bottom: 2px solid rgb(204, 204, 204); font-weight: bold;'><span style='font-weight: bolder; font-size: 20px; margin-right: 6px;'>汇全球之智 &nbsp; 明安全之道</span></p></div></div><div style='font-family: Verdana; font-size: 14px;'>&nbsp;</div></div></span></div><img src='cid:image1'>"
	content = MIMEText(email_content,'html' ,'utf-8')
	#附件路径
	msg = MIMEMultipart()
	msg.attach(content)        
	#msg['Received'] = Header("from [127.0.0.1] (unknown [183.129.153.153])by myvps2(Postfix) with ESMTPA id CBBC9101AF2")
	msg['From'] = 'zh_l@jslhbank.com'
	#msg['From'] = 'zhanghua@ccb.com'
	msg['Subject'] = Header('测试方案','utf-8').encode()
	#邮件带图片
	#file = open('3.png', 'rb')
	#img_data = file.read()
	#file.close()
	#img = MIMEImage(img_data,_subtype='octet-stream')
	#img.add_header('Content-ID', '<image1>')
	#msg.attach(img)
	#带附件
	#docfile = '保密知识考试材料.rar'
	#docpart = MIMEApplication(open(docfile, 'rb').read())
	#docpart.add_header('Content-Disposition', 'attachment', filename = docfile)
	#msg.attach(docpart)
	
	
	#print(to_addr.strip())
	msg['To'] = to_addr
	#print(msg['To'])
	
	server = smtplib.SMTP(smtp_server,25)
	#server.login(from_addr,password)
	try:
		server.sendmail(from_addr, to_addr, msg.as_string())
		print('发送成功--' + to_addr)
	except Exception as e:
		print('发送失败，再次尝试\t' + to_addr)
		print(e)
		
		#server.sendmail(from_addr, to_addr, msg.as_string())
	#print('发送邮件到'+to_addr)
	server.quit()
	
