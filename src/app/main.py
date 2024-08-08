# -*- coding: utf-8 -*-

'''
檔案說明：主流程控制
Writer：Qian
'''
"""
主功能步驟
1.偵測系統時間，晚上22點開始運作
2.從onedrive下載檔案到本地
3.檢查DealerInfo是否更新
4.執行檔案檢查
5.等待10分鐘
6.判斷對應目錄是否檔案有更新
7.onedrive下載更新檔案
8.執行檔案檢查
9.等待10分鐘
10.判斷對應目錄是否檔案有更新
11.onedrive下載更新檔案
12.執行檔案檢查
13.轉換檔案
14.結算報表
15.上傳至EFT
16.轉換完畢的檔案上傳至onedrive
"""

import os
import time
import schedule
from datetime import datetime
from Config import AppConfig
from FileInfo import CheckFileInfo
from OneDriveFile import DownloadOneDrive
from CheckFile import RecordDealerFiles, CheckFile, ClearSubRecordJson

Config = AppConfig()
Could = os.path.join(Config.OneDrivePath, Config.FolderName)
Local = os.path.join(Config.RootDir, Config.FolderName)
SkipFolder = Config.NotCopyFolder

def start_and_end_in_check():
    result = DownloadOneDrive(Could, Local, SkipFolder)
    if result:
        CheckFileInfo()
        have_submission, sub_dic, sub, resub = RecordDealerFiles(Config.TestMode)
        print(f"have_submission:{have_submission}")
        print(f"sub_dic:{sub_dic}")
        print(f"sub:{sub}")
        print(f"resub:{resub}")
        change_dic = CheckFile(have_submission, sub_dic, sub, resub)
        print(f"change_dic:{change_dic}")

def main():
    print()

if __name__ == "__main__":
    start_and_end_in_check()

# schedule.every().day.at("03:19").do(main)

# try:
#     print("start")
#     while True:    
#         schedule.run_pending()
#         time.sleep(15)
# except KeyboardInterrupt:
#     print("Scheduler stopped by user.")
