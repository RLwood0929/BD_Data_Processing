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

"""
---------------------------------------------------------------------------------------------------------------------------------
"""

import os
import smtplib
from dotenv import load_dotenv
from email.mime.text import MIMEText #內容使用
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication #附件使用
from SystemConfig import MailRule, User, Config, DealerConf, SubRecordJson

GlobalConfig = Config()
DealerConfig = DealerConf()
MailConfig = MailRule()
UserConfig = User()
MailSendConfig = SubRecordJson("Read", None)

DealerList = DealerConfig["DealerList"]
TemplateFolderPath = GlobalConfig["Default"]["MailTemplate"]

# 依據mode event回傳對應的 index
def get_mail_index(mode, dealer_index):
    flag = True
    count = 0
    if mode == "EFTConnectError":
        index = 1
    elif mode == "FileNotSub":
        index = 2
        count = MailSendConfig[f"Dealer{dealer_index}"]["Mail2"]
    elif mode == "FileReSubError":
        index = 3
        count = MailSendConfig[f"Dealer{dealer_index}"]["Mail3"]
    elif mode == "FileContentError":
        index = 4
        count = MailSendConfig[f"Dealer{dealer_index}"]["Mail4"]
    elif mode == "ChangeReport":
        index = 5
    elif mode == "ErrorReport":
        index = 6
    elif mode == "MasterFileMaintain":
        index = 7
    else:
        flag = False

# 讀取信件 html 模板
def load_template(template_name):
    template_path = os.path.join(TemplateFolderPath, template_name)
    with open(template_path, "r", encoding = "UTF-8") as html_file:
        template_content = html_file.read()
    return template_content

# 取得信件內容，html中的變量給予值
def get_mail_content(mail_index, template, mail_data):
    subject = None
    if mail_index == 1:
        # mail_data = {"DateTime": date_time}
        date_time = mail_data["DateTime"]
        mail_content = template.format(DateTime = date_time)
    elif mail_index == 2:
        # mail_data = {"FileName": file_name,"DateTime":date_time}
        file_name = mail_data["FileName"]
        date_time = mail_data["DateTime"]
        mail_content = template.format(FileName = file_name, DateTime = date_time)
    elif mail_index == 3:
        # mail_data = {"FileName":file_name, "FileType": file_type}
        file_name = mail_data["FileName"]
        file_type = mail_data["FileType"]
        mail_content = template.format(FileName = file_name, FileType = file_type)
    elif mail_index == 4:
        # mail_data = {"FileName": file_name}
        file_name = mail_data["FileName"]
        mail_content = template.format(FileName = file_name)
    elif mail_index == 5:
        # mail_data = {"FileNum" : file_num, "DataNum" : data_num, "CheckErrorNum" : check_error_num,\
        #              "ChangeErrorNum" : change_error_num, "ReportName", report_name}
        file_num = mail_data["FileNum"]
        data_num = mail_data["DataNum"]
        check_error_num = mail_data["CheckErrorNum"]
        change_error_num = mail_data["ChangeErrorNum"]
        report_name = mail_data["ReportName"]
        mail_content = template.format(FileNum = file_num,\
                                       DataNum = data_num,\
                                       CheckErrorNum = check_error_num,\
                                       ChangeErrorNum = change_error_num,\
                                       ReportName = report_name)
    elif mail_index == 6:
        # mail_data = {"ErrorReportFileName" : error_report_file_name}
        error_report_file_name = mail_data["ErrorReportFileName"]
        mail_content = template.format(ErrorReportFileName = error_report_file_name)
        subject = subject.replace("{DealerID}", dealer_id)
    elif mail_index == 7:
        # mail_data = {"DataNum":data_num, "DateTime":date_time, "OneDriveLink":one_drive_link}
        data_num = mail_data["DataNum"]
        date_time = mail_data["DateTime"]
        one_drive_link = mail_data["OneDriveLink"]
        mail_content = template.format(DataNum = data_num,\
                                       DateTime = date_time,\
                                       OneDriveLink = one_drive_link)
        subject = subject.replace("{DealerID}", dealer_id)

    return subject, mail_content


    
    return flag, count, index

# 依據 Mode Event，選定收件者、信件模板等資訊，回傳為 dict 型態
# EFTConnectError、FileNotSub、FileReSubError、FileContentError、ChangeReport、ErrorReport、MasterFileMaintain
# mail_info = {"Mode" : mode, "Subject" : subject, "Recipients" : recipient_list,\
#              "CopyRecipients" : copy_recipient_list, "MailContent" : mail_content}
def GetMailInfo(mode, dealer_id, mail_data):
    for i in range(len(DealerList)):
        if DealerList[i] == dealer_id:
            dealer_index = i + 1
            DealerInfo = DealerConfig[f"Dealer{dealer_index}"]
            dealer_mail = DealerInfo.get("Contact1Mail")\
                or DealerInfo.get("Contact2Mail")\
                or DealerInfo.get("ContactProjectMail")
            break
    
    result, mail_count, mail_index = get_mail_index(mode, dealer_index)
    if result:
        subject = MailConfig[f"Mail{mail_index}"]["Subject"]
        recipient = MailConfig[f"Mail{mail_index}"]["Recipient"]
        copy_recipient = MailConfig[f"Mail{mail_index}"]["CopyRecipient"]
        templates = MailConfig[f"Mail{mail_index}"]["Content"]
        template = load_template(templates)
        repeatedly = MailConfig[f"Mail{mail_index}"]["Repeatedly"]

        subject_new, mail_content = get_mail_content(mail_index, template, mail_data)
        if subject_new:
            subject = subject_new

        # 取得收件者 Mail
        recipient_list = []
        if "Dealer" in recipient:
            recipient_list.append(dealer_mail)
            mail_count += 1
            write_data = {f"Dealer{dealer_index}":{f"Mail{mail_index}": mail_count}}
            SubRecordJson("WriteFileStatus", write_data)
            recipient.extend(repeatedly) if mail_count >= 3 else None

        for i in range(len(UserConfig)):
            index = i + 1
            user_group = UserConfig[f"User{index}"]["Group"]
            user_mail =  UserConfig[f"User{index}"]["Mail"]
            if user_group in recipient:                
                if (mail_count >= 3) and (user_group == "BD_BA"):
                    ba_responsible = UserConfig[f"User{index}"]["ResponsibleDealerID"]
                    if ba_responsible and (dealer_id in ba_responsible):
                        ba_mail = UserConfig[f"User{index}"]["Mail"]
                        recipient_list.append(ba_mail)
                else:
                    recipient_list.append(user_mail)
        
        # 取得副本收件者 Mail
        copy_recipient_list = []
        if copy_recipient:
            for group in copy_recipient:
                for i in range(len(UserConfig)):
                    index = i + 1
                    user_group = UserConfig[f"User{index}"]["Group"]
                    user_mail =  UserConfig[f"User{index}"]["Mail"]
                    if group == user_group:
                        copy_recipient_list.append(user_mail)

        mail_info = {"Mode" : mode, "Subject" : subject, "Recipients" : recipient_list,\
                    "CopyRecipients" : copy_recipient_list, "MailContent" : mail_content}

        return mail_info

def eft_connect_error():
    print()



if __name__ == "__main__":
    dealer_id = "111"
    Mail_info = GetMailInfo("EFTConnectError", dealer_id, {"DateTime":"2024/07/29"})
    print(Mail_info)