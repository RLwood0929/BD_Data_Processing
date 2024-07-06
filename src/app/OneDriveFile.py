# -*- coding: utf-8 -*-

'''
檔案說明：OneDrive上傳或下載檔案
Writer:Qian
'''

import os
import shutil
from app.Logs import WSysLog
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
SourcePath = os.path.join(OneDrivePath, FolderName)
TargetPath = os.path.join(RootDir, FolderName)

# 確認本地 OneDrive 路徑是否存在
def DrivePathCheck():
    if os.path.exists(OneDrivePath):
        check_message = "OneDrive 路徑存在"
        WSysLog("1", "DrivePathCheck", check_message)
        return True
    else:
        check_message = "OneDrive 路徑不存在，請檢查本地 OneDrive 目錄，並且登入嘉衡公司帳戶"
        WSysLog("4", "DrivePathCheck", check_message)
        return False
    
# 建構本地目錄
def MakeLocalFolder():
    try:
        os.makedirs(TargetPath)
        make_message = f"成功建立目錄{TargetPath}"
        WSysLog("1", "MakeLocalFolder", make_message)
        return True
    except Exception as e:
        make_message = f"發生錯誤: {e}"
        WSysLog("3", "MakeLocalFolder", make_message)
        return False

# 確認本地目標路徑是否存在
def TargetPathCheck():
    if os.path.exists(TargetPath):
        check_message = "本地目標路徑存在"
        WSysLog("1", "TargetPathCheck", check_message)
        return True
    else:
        msg = "本地路徑不存在，需創建，跳轉至創建函數 MakeLocalFolder"
        WSysLog("1", "TargetPathCheck", msg)
        result = MakeLocalFolder()
        return result

# 差異同步後進行拷貝
def CopyFolders(source, destination, skip_lsit = []):
    def recursive_sync(src, dst):
        copy_flag = False
        for root, dirs, files in os.walk(src):
            # 從目錄清單中刪除要跳過的目錄
            dirs[:] = [d for d in dirs if d not in skip_lsit]

            # 子目錄拷貝
            for name in dirs:
                src_dir = os.path.join(root, name)
                dst_dir = os.path.join(dst, os.path.relpath(src_dir, src))
                try:
                    if not os.path.exists(dst_dir):
                        os.makedirs(dst_dir)
                        msg = f"成功建立目錄 {dst_dir}"
                        WSysLog("1", "CopyFolders_CopySubdirectory", msg)
                        copy_flag = True
                except OSError as e:
                    msg = f"創建目錄 Error {dst_dir}: {e}"
                    WSysLog("3", "CopyFolders_CopySubdirectory", msg)
            
            # 檔案拷貝
            for name in files:
                src_file = os.path.join(root, name)
                dst_file = os.path.join(dst, os.path.relpath(src_file, src))
                try:
                    if not os.path.exists(dst_file) or os.stat(src_file).st_mtime > os.stat(dst_file).st_mtime:
                        shutil.copy2(src_file, dst_file)
                        msg = f"拷貝 {src_file} 至 {dst_file}"
                        WSysLog("1","CopyFolders_CopyFiles", msg)
                        copy_flag = True
                except (IOError, shutil.Error) as e:
                    msg = f"拷貝 Error {src_file} 至 {dst_file}: {e}"
                    WSysLog("3","CopyFolders_CopyFiles", msg)
                
        for root, dirs, files in os.walk(dst):
            for name in dirs:
                src_dir = os.path.join(root, name)
                rel_dir = os.path.relpath(src_dir, dst)
                src_dir_full = os.path.join(src, rel_dir)
                if not os.path.exists(src_dir_full):
                    try:
                        shutil.rmtree(src_dir)
                        msg = f"刪除目錄 {src_dir}"
                        WSysLog("1", "CopyFolders_RemoveExtraFolders", msg)
                        copy_flag = True
                    except Exception as e:
                        msg = f"刪除 Error {src_dir}: {e}"
                        WSysLog("3", "CopyFolders_RemoveExtraFolders", msg)
            
            for name in files:
                src_file = os.path.join(root, name)
                rel_file = os.path.relpath(src_file, dst)
                src_file_full = os.path.join(src, rel_file)
                if not os.path.exists(src_file_full):
                    try:
                        os.remove(src_file)
                        msg = f"刪除檔案 {src_file}"
                        WSysLog("1", "CopyFolders_RemoveExtraFiles", msg)
                        copy_flag = True
                    except Exception as e:
                        msg = f"刪除{src_file} Error: {e}"
                        WSysLog("3", "CopyFolders_RemoveExtraFiles", msg) 

        if not copy_flag:
            msg = "OneDrive 與 本地目錄 無差異"
            WSysLog("1", "CopyFolders", msg)

    recursive_sync(source, destination)

# 從 OneDrive 下載檔案
def DownloadFOD(skip_list):
    if DrivePathCheck() & TargetPathCheck():
        # 冷資料應略過
        CopyFolders(SourcePath, TargetPath, skip_list)

# 上傳檔案到 OneDrive
def UploadTOD(skip_list):
    if DrivePathCheck() & TargetPathCheck():
        CopyFolders(TargetPath, SourcePath, skip_list)

if __name__ == "__main__":
    # 略過的目錄
    SkipList = []
    DownloadFOD(SkipList)