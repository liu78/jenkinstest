#coding=utf-8
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib
def send_email(x):
    smtpserver = 'smtp.163.com'
    user = 'liuchinho@163.com'
    password = '****mlgb'
    sender = 'liuchinho@163.com'
    receiver = ['360482821@qq.com']
    subject = u'脚本执行完成报告'
    msg = MIMEText(x, 'plain', 'utf-8')
    msg['From'] = 'liuchinho@163.com<liuchinho@163.com>'
    msg['To'] = ";".join(receiver)
    msg['Subject'] = Header(subject, 'utf-8')
    
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver, 25)
    smtp.login(user, password)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()
    
str = 'Done!'
send_email(str)