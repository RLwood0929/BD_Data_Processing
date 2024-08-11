# -*- coding: utf-8 -*-

'''
檔案說明：EFT雲端檔案透過FTP傳輸
Writer：Qian
'''

import os
import ftplib
from Log import WSysLog
from Mail import SendMail
from Config import AppConfig

Config = AppConfig()

# EFT連線測試
def EFTConnectTest():
    msg = "進行EFT連線測試。"
    WSysLog("1", "EFTContectTest", msg)
    try:
        ftp_server = ftplib.FTP(Config.EFTHostName)
        ftp_server.login(Config.EFTUserName, Config.EFTPassword)
        msg = "登入成功。"
        WSysLog("1", "EFTLoginTest", msg)
        ftp_server.retrlines("LIST")
        msg = "EFT連線測試成功。"
        WSysLog("1", "EFTContectTest", msg)
    except ftplib.all_errors as e:
        # 取得EFT連線時的錯誤資訊
        msg = f"登入失敗: {e}"
        WSysLog("3", "EFTLoginTest", msg)
        msg = "EFT連線測試失敗。"
        WSysLog("1", "EFTContectTest", msg)
    finally:
        # 關閉EFT雲端連線
        if ftp_server:
            ftp_server.quit()
            msg = "EFT FTP雲端連線已關閉。"
            WSysLog("1", "EFTUploadFile", msg)

# 刪除EFT測試資料
def DeleteEFTFile():
    print("刪除測試資料")
    try:
        ftp_server = ftplib.FTP(Config.EFTHostName)
        ftp_server.login(Config.EFTUserName, Config.EFTPassword)
        remote_path = os.path.join(Config.EFTDir, "README.md")
        print(remote_path)
        ftp_server.delete(remote_path)
        print("刪除成功")
    except ftplib.all_errors as e:
        print(f"An error occurred: {e}")
    finally:
        # 關閉EFT雲端連線
        if ftp_server:
            ftp_server.quit()
            print("結束 EFT 雲端 FTP Server 連線")

# 取得目標目錄底下的檔案
def get_files(folder_path):
    # 確保目錄存在
    if not os.path.exists(folder_path):
        msg = f"{folder_path} 目錄不存在。"
        WSysLog("3", "CheckChangeFolderPath" , msg)
        try:
            os.makedirs(folder_path)
            msg = f"已建立 {folder_path} 目錄。"
            WSysLog("1", "CheckChangeFolderPath" , msg)
            return None
        except Exception as e:
            msg = f"建立 {folder_path} 目錄失敗，原因{e}。"
            WSysLog("3", "CheckChangeFolderPath" , msg)
            return False
          
    # 取得 ChangeFile 目錄底下的檔案
    file_names = [file for file in os.listdir(folder_path) \
                    if os.path.isfile(os.path.join(folder_path, file))]
    
    if len(file_names) == 0:
        msg = f"{folder_path} 來源目錄底下檔案數量為 {len(file_names)}。"
        WSysLog("1", "CheckExchangeFileNum", msg)
    return file_names

# 將ChangeFile目錄底下檔案上傳至EFT雲端
def upload_file(folder_path, file_names, retry_count = 0, max_retries = Config.MaxTryRange):
    try:
        ftp_server = ftplib.FTP(Config.EFTHostName)
        ftp_server.login(Config.EFTUserName, Config.EFTPassword)
        ftp_server.cwd(Config.EFTDir)

        # 測試模式
        if not Config.TestMode:
            file_name = "README.md"
            file_path = "./README.md"
            with open(file_path, "rb") as file:
                ftp_server.storbinary(f"STOR {file_name}", file)

            remote_size = ftp_server.size(file_name)
            local_size = os.path.getsize(file_path)

            # 比對檔案大小，判斷檔案是否上傳成功
            if remote_size == local_size:
                msg = f"成功上傳測試檔案至 EFT雲端：{Config.EFTDir}。"
                WSysLog("1", "EFTUploadFile", msg)
            else:
                msg = f"上傳測試檔案至 EFT雲端：{Config.EFTDir}失敗，請重新上傳。"
                WSysLog("3", "EFTUploadFile", msg)
        
        # 一般模式
        else:
            upload_flag = True
            error_files = []
            for file_name in file_names:
                file_path = os.path.join(folder_path, file_name)
                with open(file_path, "rb") as file:
                    ftp_server.storbinary(f"STOR {file_name}", file)
                
                # 比對檔案大小，判斷檔案是否上傳成功
                remote_size = ftp_server.size(file_name)
                local_size = os.path.getsize(file_path)
                if remote_size != local_size:
                    upload_flag = False
                    error_files.append(file_name)
            if upload_flag:
                msg = f"成功上傳 {len(file_names)} 份檔案至 EFT雲端：{Config.EFTDir}。"
                WSysLog("1", "EFTUploadFile", msg)
            else:
                if retry_count < max_retries:
                    msg = f"上傳 {error_files} 至 EFT 雲端發生錯誤，系統嘗試重新上傳(嘗試次數:{retry_count + 1})。"
                    WSysLog("3", "EFTUploadFile", msg)
                    upload_file(folder_path, error_files, retry_count + 1, max_retries)
                else:
                    mail_data = {"FileName":"、".join(error_files)}
                    send_info = {"Mode" : "EFTUploadFileError", "DealerID" : None, "MailData" : mail_data, "FilesPath" : None}
                    SendMail(send_info)
                    msg = f"上傳 {error_files} 至 EFT 雲端失敗，已達最大重試次數。"
                    WSysLog("3", "EFTUpload", msg)
        return True
    
    except ftplib.all_errors as e:
        mail_data = {"DateTime":Config.SystemTime}
        send_info = {"Mode" : "EFTConnectError", "DealerID" : None, "MailData" : mail_data, "FilesPath" : None}
        SendMail(send_info)
        msg = f"上傳檔案遇到錯誤，錯誤原因{e}。"
        WSysLog("3", "EFTUploadFile", msg)
        EFTConnectTest()
        return False
    
    finally:
        # 關閉EFT雲端連線
        if "ftp_server" in locals():
            ftp_server.quit()
            msg = "EFT FTP雲端連線已關閉。"
            WSysLog("1", "EFTUploadFile", msg)

# 檔上傳至EFT雲端主程式
def EFTUploadFile():
    folder_path = Config.ChangeFolderPath
    file_names = get_files(folder_path)
    if file_names:
        upload_file(folder_path, file_names)
    
if __name__ == "__main__":
    # EFTConnectTest()
    EFTUploadFile()
    # DeleteEFTFile()