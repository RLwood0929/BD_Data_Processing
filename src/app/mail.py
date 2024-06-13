# -*- coding: utf-8 -*-

'''
檔案說明：套用對應主旨、信件模板及寄件者，
使用指定的mail帳戶發送信件
Writer：Qian
'''

import os
import smtplib
from dotenv import load_dotenv
from email.mime.text import MIMEText #內容使用
from email.mime.multipart import MIMEMultipart
#from email.mime.application import MIMEApplication #附件使用

def load_template(template_name):
    template_path = os.path.join("src", "templates", template_name)
    with open(template_path, "r", encoding="utf-8") as file:
        template_content = file.read()
    return template_content

def send(Subject, Recipients, HtmlContent):
    load_dotenv()
    sender = os.getenv("Sender")
    password = os.getenv("SenderPassword")
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = Recipients
    msg["Subject"] = Subject
    html_part = MIMEText(HtmlContent, "html")
    msg.attach(html_part)

    with smtplib.SMTP(host=os.getenv("MailHost"), port=os.getenv("Mailport")) as smtp:
        try:
            smtp.ehlo() #驗證smtp伺服器
            smtp.starttls() #建立加密傳輸
            smtp.login(sender, password)
            smtp.send_message(msg)
            print("send_success")
            return ("send_success")
        except Exception as e:
            print("Error massage: ", e)
            return(str(e))

def test():
    Subject = "BD資料清洗專案信件測試(信件為系統自動發送請勿回復)"
    Recipients = "richardwu@coign.com.tw"
    html_template = load_template("Change_Report.html")
    html_content = html_template
    send(Subject,Recipients,html_content)

if __name__== "__main__":
    test()