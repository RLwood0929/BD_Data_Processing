# -*- coding: utf-8 -*-

'''
檔案說明：OneDrive上傳或下載檔案
Writer:Qian
'''

import os
# import shutil
from SystemConfig import Config

GlobalConfig = Config()

WinUser = GlobalConfig["App"]["WinUser"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
OneDrivePath = GlobalConfig["App"]["OneDrivePath"] if GlobalConfig["App"]["OneDrivePath"] \
    else GlobalConfig["Default"]["OneDrivePath"]
OneDrivePath = OneDrivePath.replace("{username}", WinUser)
TargetPath = os.path.join(RootDir, FolderName)

# 確認本地 OneDrive 路徑是否存在
def DrivePathCheck():
    if os.path.exists(OneDrivePath):
        check_message = "OneDrive 路徑存在"
        return True, check_message
    else:
        check_message = "OneDrive 路徑不存在，請檢查本地 OneDrive 目錄，並且登入嘉衡公司帳戶"
        return False, check_message
    
# 建構本地目錄
def MakeLocalFolder():
    try:
        os.makedirs(TargetPath)
        make_message = f"成功建立目錄{TargetPath}"
        return True, make_message
    except Exception as e:
        make_message = f"An error occurred: {e}"
        return False, make_message

# 確認本地目標路徑是否存在
def TargetPathCheck():
    if os.path.exists(TargetPath):
        check_message = "本地目標路徑存在"
        return True, check_message
    else:
        result, check_message = MakeLocalFolder()
        return result, check_message
        

'''
# 從 OneDrive 下載檔案 (Download File From OneDrive)
def DownloadFileFOD():
    
    # 抓取目錄路徑
    load_dotenv()
    RootDir = os.getenv("RootDir")
    ODPath = os.getenv("OneDrivePath")

    TargetDir = os.path.join(RootDir, BDLocalPath)

    if os.path.exists(TargetDir):
        print("TargetDir is exist.")
    else:
        MakeTargetDir(RootDir, BDLocalPath)

    # 本地OneDrive目錄
    ODDir = os.path.join(ODPath,"BD_DataProcessing")
    print(ODDir)
    SorceDir = ODDir
    CopyFile(SorceDir,TargetDir)

# 上傳檔案至 OneDrive (Upload File To OneDrive)
def UploadFileTOD():

    load_dotenv()
    RootDir = os.getenv("RootDir")
    ODPath = os.getenv("OneDrivePath")

    TargetDir = os.path.join(RootDir, BDLocalPath)

# 拷貝檔案
def CopyFile(Src, Dst):
    try:
        for item in os.listdir(Src):
            s = os.path.join(Src, item)
            d = os.path.join(Dst, item)

            if os.path.isdir(s):
                CopyFile(s, d)
            else:
                # 如果是文件則直接拷貝
                shutil.copy2(s, d)
        print(f"所有文件和資料夾已成功拷貝至 {Dst}")

        if not os.listdir(Src):
            if not os.path.exists(Dst):
                os.makedirs(Dst)

    except Exception as e:
        print(f"發生錯誤: {e}")
'''

if __name__ == "__main__":
    result,message = TargetPathCheck()
    print(result)
    print(message)