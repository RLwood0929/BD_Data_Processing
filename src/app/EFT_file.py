# -*- coding: utf-8 -*-

'''
檔案說明：EFT雲端檔案透過FTP傳輸
Writer：Qian
'''

import os
import ftplib
from dotenv import load_dotenv
from SystemConfig import Config

# 從json檔案取得config
GlobalConfig = Config()
EFTHostName = GlobalConfig["EFT"]["HostName"]
EFTDir = GlobalConfig["EFT"]["Dir"]

# 從env檔案取得環境變數
load_dotenv()
username = os.getenv("EFTUser")
password = os.getenv("EFTPwd")

# EFT連線測試
def EFTConnect():
    try:
        ftp = ftplib.FTP(EFTHostName)
        ftp.login(user=username, passwd=password)
        print("登入成功")
        ftp.retrlines("LIST")
    except ftplib.all_errors as e:
        # 捕获所有FTP相关的错误
        print(f"登入失敗: {e}")
    finally:
        # 确保连接被关闭
        if ftp:
            ftp.quit()

# 透過FTP上傳檔案至EFT雲端
def EFTUploadFile(LocalPath, FileName):
    try:
        ftp = ftplib.FTP(EFTHostName)
        ftp.login(user=username, passwd=password)
        ftp.cwd(EFTDir)
        FilePath = os.path.join(LocalPath,FileName)
        with open(FilePath,"rb") as file:
            ftp.storbinary(f"STOR {FileName}", file)
        print("成功上傳")
    except ftplib.all_errors as e:
        print(f"Error: {e}")

# 透過FTP下載檔案至本地
def EFTDownloadFile(LocalPath, FileName):
    try:
        ftp = ftplib.FTP(EFTHostName)
        ftp.login(user=username, passwd=password)
        ftp.cwd(EFTDir)
        FilePath = os.path.join(LocalPath, FileName)
        with open(FilePath,"wb") as file:
            ftp.retrbinary(f'RETR {FileName}', file.write)
        print("成功下載")
    except ftplib.all_errors as e:
        print(f"Error: {e}")

if __name__ == "__main__" :
    localPath = "datas"
    file_name = "test.xlsx"
    EFTDownloadFile(localPath, file_name)