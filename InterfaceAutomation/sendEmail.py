#! /usr/bin/env python
# coding=utf-8

import smtplib
from email.mime.text import MIMEText
from email.header import Header
# import sys
# reload(sys)
# sys.setdefaultencoding( "utf-8" )

# 第三方 SMTP 服务
# mail_host="smtp.exmail.qq.com"  #设置服务器
mail_host="smtp.163.com"  #设置服务器
mail_user="interfacetest11@163.com"    #用户名
mail_pass="Welcome1"   #口令

def SendMail(From, To ,Title ,mail_msg):
    #@
    #    From：发件人
    #    To：收件人
    #    Title：邮件标题
    #    mail_msg：邮件内容（可以是html，或文本）
    
    message = MIMEText(mail_msg, 'html', 'utf-8')
    message['From'] = Header(From, 'utf-8')
    message['To'] = Header(To, 'utf-8')

#     subject = 'Python SMTP 邮件测试'
    message['Subject'] = Header(Title, 'utf-8')
    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
        # smtpObj = smtplib.SMTP_SSL(mail_host, 465)  # 启用SSL发信，端口一般是465
        print("+++++++++++++++++++++++++")
        smtpObj.login(mail_user,mail_pass)
        print("================")
        smtpObj.sendmail(From, To, message.as_string())
        # smtpObj.sendmail("interfacetest11@163.com", "interfacetest11@163.com", message.as_string())
        # print ("邮件发送成功").decode('utf-8')
        print("邮件发送成功")
    except smtplib.SMTPException:
        # print ("Error: 无法发送邮件").decode('utf-8')
        print("Error: 无法发送邮件")

if __name__ == '__main__':
    mail_msg = \
    """
        <p>内容：Python 邮件发送测试...</p>
        <p><a href="http://www.baidu.com">官网链接</a></p>
    """
#     with open('sample.html', 'r') as newF:
#         mail_msg = newF.read();
#     newF.close();
    SendMail(u"interfacetest","interfacetest11@163.com","我来测试一下发邮件的方法" ,mail_msg)
