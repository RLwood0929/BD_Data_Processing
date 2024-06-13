# -*- coding: utf-8 -*-

'''
檔案說明：OneDrive上傳或下載檔案
Writer:Qian
'''

import os
import shutil
from dotenv import load_dotenv

BDLocalPath = "BD_DataProcessing"
FolderPath = ["00_System\\Log\\01_Success_Log","00_System\\Log\\02_Error_Log","01_BD\\01_MainFile","01_BD\\02_Report","02_Dealer"]

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

# 初始建立資料夾
def MakeTargetDir(RootDir, NextDir):
    TargetDir = os.path.join(RootDir, NextDir)
    os.makedirs(TargetDir)
    for i in range(len(FolderPath)):
        folder = os.path.join(TargetDir,FolderPath[i])
        os.makedirs(folder)

if __name__ == "__main__":
    DownloadFileFOD()
