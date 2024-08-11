# -*- coding: utf-8 -*-

'''
檔案說明：主流程控制
Writer：Qian
'''
"""
系統自動運行步驟
01. 系統開始
02. 執行OneDrive目錄同步至本地
03. 執行OneDrive檔案同步至本地
04. 記錄經銷商目錄底下有檔案的經銷商
05. 將Config檔案對比本地目錄，若雲端變更時間版本較新，從雲端更新至本地
06. 進行BA目錄底下檔案檢查，對比檔案json紀錄看是否有更新
07. 經銷商檔案繳交狀況記錄 > mail缺繳通知
08. 經銷商檔案檢查 > mail 檔案錯誤通知
09. 第一階段結束、第二階段結束
10. 檢查出錯誤的檔案歸位
11. 取得前幾次運作系統檢查完的結果
12. 經銷商檔案傳換 > mail error report
13. 合併庫存檔案生成txt檔
14. 轉換結果上傳至EFT雲端(3次上傳測試)
15. 上傳完後將轉換完畢的檔案規檔
16. 將檢查正確檔案規檔
17. 系統產生總結記錄表
18. 依據檔案缺繳 rawdata(今天以前的) ，mail催繳通知，(無檢查，表頭正確；內容正確)
19. 清理檔案繳交json
20. 將系統Config檔案搬到本地目錄
21. 上傳回OneDrive
22. 清空雲端經銷商上傳目錄底下的檔案
23.  清空本地檔案
24. 系統結束
"""

import os, sys
import time
import schedule
from Config import AppConfig
from Log import WSysLog
from SystemConfig import SubRecordJson
from FileInfo import CheckFileInfo, ConfigFile, ConfigFileToCould
from OneDriveFile import DownloadOneDrive, UploadOneDrive, ClearLocal
from CheckFile import RecordDealerFiles, CheckFile, MoveCheckErrorFile, MoveCheckFile, ClearSubRecordJson
from Mapping import Changing, MergeInventoryFile, FileArchiving
from EFTFile import EFTUploadFile
from RecordTable import Statistics

Config = AppConfig()

def system_work_flow(half_flag = False):
    Could = os.path.join(Config.OneDrivePath, Config.FolderName)
    Local = os.path.join(Config.RootDir, Config.FolderName)
    SkipFolder = Config.NotCopyFolder
    try:
        print("=== System Start ===")
        print("--Running DownloadOneDrive--")
        result, DealerList = DownloadOneDrive(Could, Local, SkipFolder)
        print("Result:")
        print(f"\tDealerList:{DealerList}")
        if not result:
            raise SystemError("DownloadOneDrive區塊發生錯誤，請查閱log紀錄。")
        print("--End DownloadOneDrive--")
        print("--Running ConfigFile--")
        result = ConfigFile()
        if not result:
            raise FileNotFoundError("ConfigFile區塊發生錯誤，請查閱log紀錄。")
        print("--End ConfigFile--")
        print("--Running CheckFileInfo--")
        CheckFileInfo()
        print("--End CheckFileInfo--")
        print("--Running RecordDealerFiles--")
        HaveSubmission, SubDic, Sub, ReSub = RecordDealerFiles("AutoRun", DealerList)
        print("Result:")
        print(f"\tHaveSubmission:{HaveSubmission}")
        print(f"\tSubDic:{SubDic}")
        print(f"\tSub:{Sub}")
        print(f"\tRsSub:{ReSub}")
        print("--End RecordDealerFiles--")
        print("--Running CheckFile--")
        ChangeDic =  CheckFile(HaveSubmission, SubDic, Sub, ReSub)
        print("Result:")
        print(f"\tChangeDic:{ChangeDic}")
        if ChangeDic:
            old_data = SubRecordJson("ReadChangeDic")
            ChangeDic.update(old_data)
            msg = SubRecordJson("WriteChangeDic", ChangeDic)
            WSysLog("1", "SubRecordJson", msg)
        print("--End CheckFile--")
        if half_flag:
            print("=== System Break ===")
            return
        print("--Running MoveCheckErrorFile--")
        MoveCheckErrorFile()
        print("--End MoveCheckErrorFile--")
        print("--Running SubRecordJson--")
        ChangeDic = SubRecordJson("ReadChangeDic", None)
        print("Result:")
        print(f"\tChangeDic:{ChangeDic}")
        print("--End SubRecordJson--")
        if ChangeDic is None:
            print("No File Need To Change.")
        else:
            print("--Running Changing--")
            Changing(ChangeDic)
            print("--End Changing--")
            print("--Running MargeInventory--")
            MergeInventoryFile()
            print("--End MargeInventory--")
            if Config.TestMode:
                print("--Running EFTUploadFile--")
                EFTUploadFile()
                print("--End EFTUploadFile--")
            print("--Running FileArchiving--")
            FileArchiving()
            print("--End FileArchiving--")
        print("--Running MoveCheckFile--")
        MoveCheckFile()
        print("--End MoveCheckFile--")
        print("--Running Statistics--")
        Statistics()
        print("--End Statistics--")
        print("--Running SendNotSubMail--") 
        # 
        print("--End SendNotSubMail--")
        print("--Running ClearSubRecordJson--")
        ClearSubRecordJson()
        print("--End ClearSubRecordJson--")
        print("--Running ConfigFileToCould--")
        ConfigFileToCould()
        print("--End ConfigFileToCould--")
        print("--Running UploadOneDrive--")
        UploadOneDrive(Local, Could)
        print("--End UploadOneDrive--")
        print("--Running ClearLocal--")
        ClearLocal(Local)
        print("--End ClearLocal--")
        print("=== System End ===")
        if Config.TestMode:
            sys.exit()
    except Exception as e:
        msg = f"系統自動運作時發生錯誤。錯誤原因： {e}。"
        print(msg)

schedule.every().day.at("22:00").do(system_work_flow, half_flag = True)
schedule.every().day.at("22:15").do(system_work_flow, half_flag = True)
schedule.every().day.at("22:30").do(system_work_flow, half_flag = False)

def main():
    # system_work_flow(True)
    try:
        print("System Auto Run Start.")
        while True:    
            schedule.run_pending()
            time.sleep(10)
    except KeyboardInterrupt:
        print("Scheduler stopped by user.")

if __name__ == "__main__":
    main()