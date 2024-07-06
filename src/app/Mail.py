# -*- coding: utf-8 -*-

'''
檔案說明：套用對應主旨、信件模板及寄件者，
使用指定的mail帳戶發送信件
Writer：Qian
'''

import os
import smtplib
from dotenv import load_dotenv
from SystemConfig import Config
from email.mime.text import MIMEText #內容使用
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication #附件使用

GlobalConfig = Config()

# 載入信件模板
def load_template(template_name):
    template_path = os.path.join("src", "templates", template_name)
    with open(template_path, "r", encoding="utf-8") as file:
        template_content = file.read()
    return template_content

# 發送信件
def send(Subject, Recipients, CcEmail, HtmlContent, FilePath):
    load_dotenv()
    sender = os.getenv("Sender")
    password = os.getenv("SenderPassword")
    msg = MIMEMultipart()
    # 設定郵件項目
    msg["From"] = sender
    msg["To"] = ", ".join(Recipients)
    msg["Cc"] = ", ".join(CcEmail)
    msg["Subject"] = Subject
    html_part = MIMEText(HtmlContent, "html")
    msg.attach(html_part)

    # 設定郵件附件
    if FilePath:
        for file_path in FilePath:
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    file_name = os.path.basename(file_path)
                    file_extension = os.path.splitext(file_name)[1]
                    mime_subtype = file_extension[1:]
                    attachment = MIMEApplication(f.read(),_subtype=mime_subtype)
                    attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
                    msg.attach(attachment)
            else:
                print(f"File not found: {file_path}")
    else:
        print("No files to attach")

    MailHost = GlobalConfig["Mail"]["Host"]
    SMTPPort = GlobalConfig["Mail"]["SMTPPort"]

    # 發送信件
    with smtplib.SMTP(host=MailHost, port=SMTPPort) as smtp:
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

# 寄信測試
def send_test():
    Subject = "BD資料清洗專案信件測試(信件為系統自動發送請勿回復)"
    Recipients = ["richardwu@coign.com.tw"]
    CcEmail = ["wood123487@gmail.com"]
    html_template = load_template("Change_Report.html")
    html_content = html_template
    file_path = ["./README.md"]
    send(Subject,Recipients,CcEmail,html_content,file_path)

if __name__== "__main__":
    # monitor_email()
    send_test()