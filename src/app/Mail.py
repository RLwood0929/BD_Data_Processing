# -*- coding: utf-8 -*-

'''
檔案說明：套用對應主旨、信件模板及寄件者，
使用指定的mail帳戶發送信件
Writer：Qian
'''

"""
mode：EFTConnectError
mail_data = {"DateTime": date_time}

mode：FileNotSub
mail_data = {"FileName": file_name,"DateTime":date_time}

mode：FileReSubError
mail_data = {"FileName":file_name, "SubFile": sub_file_name}

mode：FileContentError
mail_data = {"FileName": file_name}
有附件

mode：ChangeReport
mail_data = {"FileNum" : file_num, "DataNum" : data_num, "CheckErrorNum" : check_error_num,
            "ChangeErrorNum" : change_error_num, "ReportName": report_name}
有附件

mode：ErrorReport 
mail_data = {"ErrorReportFileName" : error_report_file_name}
有附件

mode：MasterFileMaintain ##
mail_data = {"DataNum":data_num, "DateTime":date_time, "OneDriveLink":one_drive_link}

mode：EFTUploadFileError
mail_data = {"FileName" : file_name}

mode：MasterFileError ##
mail_data = {"MasterFileName": file_name}

"""

import os
import smtplib
import mimetypes
from Log import WSysLog
from Config import AppConfig
from email.mime.text import MIMEText #內容使用
from SystemConfig import SubRecordJson
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication #附件使用

# IMAP使用ssl驗證，未加密port號143、加密port號993

Config = AppConfig()

# 依據mode event回傳對應的 index
def get_mail_index(mode, dealer_index):
    record_data = SubRecordJson("Read", None)
    flag = True
    count = 0
    if mode == "EFTConnectError":
        index = 1
    elif mode == "FileNotSub":
        index = 2
        count = record_data[f"Dealer{dealer_index}"]["Mail2"]
    elif mode == "FileReSubError":
        index = 3
        count = record_data[f"Dealer{dealer_index}"]["Mail3"]
    elif mode == "FileContentError":
        index = 4
        count = record_data[f"Dealer{dealer_index}"]["Mail4"]
    elif mode == "ChangeReport":
        index = 5
    elif mode == "ErrorReport":
        index = 6
    elif mode == "MasterFileMaintain":
        index = 7
    elif mode == "EFTUploadFileError":
        index = 8
    elif mode == "MasterFileError":
        index = 9
    else:
        flag = False
        index = 0
    return flag, index, count

# 讀取信件 html 模板
def load_template(template_name):
    template_path = os.path.join(Config.TemplateFolderPath, template_name)
    with open(template_path, "r", encoding = "UTF-8") as html_file:
        template_content = html_file.read()
    return template_content

# 依據 Mode Event，選定收件者、信件模板等資訊，回傳為 dict 型態
# EFTConnectError、FileNotSub、FileReSubError、FileContentError、ChangeReport、ErrorReport、MasterFileMaintain
# mail_info = {"Mode" : mode, "Subject" : subject, "Recipients" : recipient_list,\
#              "CopyRecipients" : copy_recipient_list, "MailContent" : mail_content}
def GetMailInfo(mode, dealer_id, mail_data):
    if dealer_id is None:
        dealer_index = None
    else:
        for i in range(len(Config.DealerList)):
            if Config.DealerList[i] == dealer_id:
                dealer_index = i + 1
                DealerInfo = Config.DealerConfig[f"Dealer{dealer_index}"]
                dealer_mail = DealerInfo.get("Contact1Mail")\
                    or DealerInfo.get("Contact2Mail")\
                    or DealerInfo.get("ContactProjectMail")
                break

    result, mail_index, mail_count = get_mail_index(mode, dealer_index)
    if result:
        subject = Config.MailConfig[f"Mail{mail_index}"]["Subject"]
        recipient = Config.MailConfig[f"Mail{mail_index}"]["Recipient"]
        copy_recipient = Config.MailConfig[f"Mail{mail_index}"]["CopyRecipient"]
        templates = Config.MailConfig[f"Mail{mail_index}"]["Content"]
        template = load_template(templates)
        repeatedly = Config.MailConfig[f"Mail{mail_index}"]["Repeatedly"]

        # 取得信件內容，html中的變量給予值
        if mail_index == 1:
            # mail_data = {"DateTime": date_time}
            date_time = mail_data["DateTime"]
            mail_content = template.format(DateTime = date_time)

        elif mail_index == 2:
            # mail_data = {"FileName": file_name,"DateTime":date_time}
            file_name = mail_data["FileName"]
            file_name_en = []
            for name in file_name:
                if name == "銷售檔案":
                    file_name_en.append("Sale File")
                elif name == "庫存檔案":
                    file_name_en.append("Inventory File")
            date_time = mail_data["DateTime"]
            mail_content = template.format(FileName = file_name, FileNameEn = "、".join(file_name_en), DateTime = date_time)

        elif mail_index == 3:
            # mail_data = {"FileName":file_name, "SubFile": sub_file}
            file_name = mail_data["FileName"]
            sub_file = mail_data["SubFile"]
            mail_content = template.format(FileName = file_name, SubFile = sub_file)

        elif mail_index == 4:
            # mail_data = {"FileName": file_name}
            file_name = mail_data["FileName"]
            mail_content = template.format(FileName = file_name)

        elif mail_index == 5:
            # mail_data = {"FileNum" : file_num, "DataNum" : data_num, "CheckErrorNum" : check_error_num,\
            #              "ChangeErrorNum" : change_error_num, "ReportName": report_name}
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
        elif mail_index == 8:
            # mail_data = {"FileName":file_name}
            file_name = mail_data["FileName"]
            mail_content = template.format(FileName = file_name)

        elif mail_index == 9:
            # mail_data = {"MasterFileName": file_name}
            file_name = mail_data["MasterFileName"]
            mail_content = template.format(MasterFileName = file_name)
            subject = subject.replace("{DealerID}", dealer_id)

        # 取得收件者 Mail
        recipient_list = []
        if "Dealer" in recipient:
            recipient_list.append(dealer_mail)
            mail_count += 1
            write_data = {f"Dealer{dealer_index}":{f"Mail{mail_index}": mail_count}}
            SubRecordJson("WriteFileStatus", write_data)
            if mail_count >= 3:
                recipient.extend(repeatedly)

        for i in range(len(Config.UserConfig)):
            index = i + 1
            user_group = Config.UserConfig[f"User{index}"]["Group"]
            user_mail =  Config.UserConfig[f"User{index}"]["Mail"]
            if user_group in recipient:                
                if (mail_count >= 3) and (user_group == "BD_BA"):
                    ba_responsible = Config.UserConfig[f"User{index}"]["ResponsibleDealerID"]
                    if ba_responsible and (dealer_id in ba_responsible):
                        ba_mail = Config.UserConfig[f"User{index}"]["Mail"]
                        recipient_list.append(ba_mail)
                else:
                    recipient_list.append(user_mail)
        
        # 取得副本收件者 Mail
        copy_recipient_list = []
        if copy_recipient:
            for group in copy_recipient:
                for i in range(len(Config.UserConfig)):
                    index = i + 1
                    user_group = Config.UserConfig[f"User{index}"]["Group"]
                    user_mail =  Config.UserConfig[f"User{index}"]["Mail"]
                    if group == user_group:
                        copy_recipient_list.append(user_mail)

        mail_info = {"Mode" : mode, "Subject" : subject, "Recipients" : recipient_list,\
                    "CopyRecipients" : copy_recipient_list, "MailContent" : mail_content}

        return mail_info

# 使用SMTP寄送郵件
def send(mail):
    try:
        with smtplib.SMTP(Config.SMTPHost, Config.SMTPPort) as smtp:
            smtp.ehlo() #驗證smtp伺服器
            smtp.starttls() #建立加密傳輸
            smtp.login(Config.EmailSender, Config.EmailPassword)
            smtp.send_message(mail)
            return True, "send_success"
    except Exception as e:
        print("Error massage: ", e)
        return False, e

# 撰寫信件內容
def WriteMail(subject, recipients, copy_recipients, mail_content, files_path):
    #設定郵件資訊
    mail = MIMEMultipart()
    mail["From"] = Config.EmailSender
    mail["To"] = ", ".join(recipients)
    mail["Cc"] = ", ".join(copy_recipients)
    mail["Subject"] = subject
    content = MIMEText(mail_content, "html")
    mail.attach(content)

    # 添加附件進郵件中
    if files_path:
        for file_path in files_path:
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'rb') as f:
                        file_name = os.path.basename(file_path)
                        ctype, encoding = mimetypes.guess_type(file_path)
                        if ctype is None or encoding is not None:
                            ctype = 'application/octet-stream'
                        _, subtype = ctype.split('/', 1)
                        attachment = MIMEApplication(f.read(),_subtype = subtype)
                        attachment.add_header('Content-Disposition', 'attachment', filename = file_name)
                        mail.attach(attachment)
                except Exception as e:
                    msg = f"讀取檔案 {file_path} 時發生未知錯誤，錯誤原因：{e}。"
                    WSysLog("2", "AddAttachment", msg)
            else:
                msg = f"{file_path} 路徑不存在。"
                WSysLog("2", "AddAttachment", msg)

    result, msg = send(mail)
    if result:
        msg = f"主旨：{subject} 信件寄送成功。"
        WSysLog("1", "SendMail", msg)
    else:
        msg = f"主旨：{subject} 信件寄送失敗，錯誤原因{msg}。"
        WSysLog("2", "SendMail", msg)

# 寄送信件主程式
def SendMail(send_info):
    mode = send_info["Mode"]
    dealer_id = send_info["DealerID"]
    mail_data = send_info["MailData"]
    files_path = send_info["FilesPath"]
    mail_info = GetMailInfo(mode, dealer_id, mail_data)
    subject = mail_info["Subject"]

    # 測試模式
    if Config.TestMode:
        recipients = ["richardwu@coign.com.tw"]
        copy_recipients = []
    else:
        recipients = mail_info["Recipients"]
        copy_recipients = mail_info["CopyRecipients"]
    
    mail_content = mail_info["MailContent"]
    WriteMail(subject, recipients, copy_recipients, mail_content, files_path)

if __name__ == "__main__":
    MailData = {"DataNum":50, "DateTime":"2024/08/01", "OneDriveLink":"one_drive_link"}
    FilesPath = []
    test_data = {"Mode" : "MasterFileMaintain", "DealerID" : "111", "MailData" : MailData, "FilesPath" : FilesPath}
    SendMail(test_data)