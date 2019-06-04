#程序编制：ypcgamelife
#QQ:598502990
#邮箱:598502990@qq.com
#有问题或交流都可以联系。2019.6.4
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header

import xlrd
from datetime import date,datetime

import os
import os.path


#从excel中取邮件地址、附件等信息
mypath=os.getcwd()
file = '群发邮件.xls'
wb = xlrd.open_workbook(filename=file)#打开文件
#print(wb.sheet_names())#获取所有表格名字
sheet2 = wb.sheet_by_index(1)#通过索引获取表格
smtpserver=sheet2.cell(1,0).value
smtpuser=sheet2.cell(1,1).value
smtppass=sheet2.cell(1,2).value
print(smtpserver)
print(smtpuser)
print(smtppass)
sheet1 = wb.sheet_by_index(0)#通过索引获取表格
#for mailrow in range(1,2):
mailallrow=sheet1.nrows
print('mailallrows=',mailallrow)
for mailrow in range(1,mailallrow):
    sender=sheet1.cell(mailrow,0).value
    print(sender)
    if sender<'0':
        break
    recever=sheet1.cell(mailrow,1).value
    print(recever)
    message=MIMEMultipart()
    message['From']=(sender)
    message['To'] = (recever)
    subject=sheet1.cell(mailrow,2).value
    print(subject)
    message['Subject']=Header(subject)
    mailcontect=sheet1.cell(mailrow,3).value
    print(mailcontect)
    message.attach(MIMEText(mailcontect,'plain'))
    #构造附件
    attfile1=sheet1.cell(mailrow,4).value
    attfile2=sheet1.cell(mailrow,5).value
    attfile3=sheet1.cell(mailrow, 6).value
    print(attfile1+'=='+attfile2+'=='+attfile3)
    if attfile1>'0':
       mypath = os.path.split(attfile1)
       att1=MIMEText(open(attfile1,'rb').read(),'base64','utf-8')
       att1["Content-Type"]='application/octet-stream'
       #att1["Content-Disposition"]='attachment;filename=attfile1'
       att1.add_header('Content-Disposition','attachment',filename=mypath[1])
       message.attach(att1)
    if attfile2>'0':
       mypath = os.path.split(attfile2)
       att2=MIMEText(open(attfile2,'rb').read(),'base64','utf-8')
       att2["Content-Type"]='application/octet-stream'
       att2.add_header('Content-Disposition', 'attachment', filename=mypath[1])
       #att2["Content-Disposition"]='attachment;filename=attfile2'
       message.attach(att2)
    if attfile3>'0':
       mypath = os.path.split(attfile3)
       att3=MIMEText(open(attfile3,'rb').read(),'base64','utf-8')
       att3["Content-Type"]='application/octet-stream'
       #att3["Content-Disposition"]='attachment;filename=attfile1'
       att3.add_header('Content-Disposition', 'attachment', filename=mypath[1])
       message.attach(att3)

    #msg=MIMEMultipart('mixed')
    smtp = smtplib.SMTP()
    #smtp=smtplib.SMTP_SSL()
    #smtp.set_debuglevel(1)
    #smtp.ehlo()
    smtp.connect(smtpserver,25)
    smtp.ehlo()  # 向Gamil发送SMTP 'ehlo' 命令
    #smtp.starttls()
    smtp.login(smtpuser, smtppass)
    smtp.sendmail(sender,recever, message.as_string())
    smtp.quit()
    print('发送成功',recever)
print('共发送邮件',str(mailrow),'封.')
