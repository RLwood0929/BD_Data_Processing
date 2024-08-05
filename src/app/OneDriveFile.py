# -*- coding: utf-8 -*-

'''
檔案說明：OneDrive上傳或下載檔案
Writer:Qian
'''

import os
import shutil
from Log import WSysLog
from SystemConfig import Config

GlobalConfig = Config()

WinUser = GlobalConfig["App"]["WinUser"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["DirTree"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
OneDrivePath = GlobalConfig["App"]["Path"]
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

"""
===================================================================================
"""

import os, shutil
from Log import WSysLog
from Config import AppConfig
from SystemConfig import WrightWinUser, WrightOneDrivePath

Config = AppConfig()

# 取得使用者名稱，寫入config
def GetUserName():
    try:
        username = os.getlogin()
        WinUser = Config.WinUser
        if username != WinUser:
            msg = "系統使用者名稱，與紀錄不符，需更新紀錄檔。"
            WSysLog("1", "GetUserName", msg)
            msg = WrightWinUser(username)
            WSysLog("1", "WindowsUserNameUpdate", msg)
    except OSError as e:
        msg = f"取得使用者名稱時錯誤，錯誤訊息：{e}。"
        WSysLog("3", "GetUserName", msg)
    except AttributeError as e:
        msg = f"屬性錯誤，錯誤訊息：{e}。"
        WSysLog("3", "GetUserName", msg)
    except Exception as e:
        msg = f"例外錯誤，錯誤訊息：{e}。"
        WSysLog("3", "GetUserName", msg)

# 檢查OneDrive目錄在本地是否存在
def check_onedrive_path():
    if not os.path.exists(Config.OneDrivePath):
        msg = f"{Config.OneDrivePath}目錄不存在，請檢查本地 OneDrive 目錄，並且登入嘉衡公司帳戶。"
        WSysLog("3", "CheckOenDrivePath", msg)
        GetUserName()

        # 歷遍系統C槽目錄，找尋OneDrive資料夾
        for dirs, folders, _ in os.walk(Config.SystemRoot):
            if Config.OneDeiveFolder in folders:
                one_drive_path = os.path.join(dirs, Config.OneDeiveFolder)
                break
        if Config.OneDrivePath != one_drive_path:
            msg = WrightOneDrivePath(one_drive_path)
            WSysLog("1", "CheckOenDrivePath", msg)
    else:
        msg = "確認OneDrive目錄存在。"
        WSysLog("1", "CheckOneDrivePath", msg)

# 從OneDrive雲端抓取目錄結構至本地
def sync_folder(source_dir, target_dir, skip_list = []):
    sync_flag = False
    
    # 歷遍來源目錄
    for root, dirs, files in os.walk(source_dir):
        # 從目錄清單中刪除要跳過的目錄
        dirs[:] = [d for d in dirs if d not in skip_list]

        # 拷貝子目錄
        for folder in dirs:
            source_folder = os.path.join(root, folder)
            target_folder = os.path.join(target_dir, os.path.relpath(source_folder, source_dir))
            try:
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                    msg = f"建立目錄 {target_folder}。"
                    WSysLog("1", "SyncFolder_MakeFolder", msg)
                    sync_flag = True
            except OSError as e:
                msg = f"創建目錄錯誤 {target_folder} : {e}。"
                WSysLog("3", "SyncFolder_MakeFolder", msg)

        # 拷貝檔案
    #     for file in files:
    #         source_file = os.path.join(root, file)
    #         target_file = os.path.join(target_dir, os.path.relpath(source_file, source_dir))
    #         try:
    #             if not os.path.exists(target_file) or os.stat(source_file).st_mtime > os.stat(target_file).st_mtime:
    #                 shutil.copy2(source_file, target_file)
    #                 msg = f"拷貝檔案 {source_file} 至 {target_file}。"
    #                 WSysLog("1", "SyncFolder_CopyFile", msg)
    #                 sync_flag = True
    #         except (IOError, shutil.Error) as e:
    #             msg = f"拷貝 {source_file} 至 {target_file} 錯誤，錯誤原因： {e}。"
    #             WSysLog("3", "SyncFolder_CopyFile", msg)

    # # 歷遍目標目錄
    # for root, dirs, files in os.walk(target_dir):
    #     # 刪除差異子目錄
    #     for folder in dirs:
    #         target_folder = os.path.join(root, folder)
    #         print(f"source_folder:{source_folder}")
    #         source_folder_full = os.path.join(source_dir, os.path.join(source_folder, target_dir))
    #         print(f"source_folder_full:{source_folder_full}")
    #         if not os.path.exists(source_folder_full):
    #             try:
    #                 shutil.rmtree(target_folder)
    #                 msg = f"刪除目錄 {target_folder}。"
    #                 WSysLog("1", "SyncFolder_RemoveExtraFolder", msg)
    #                 sync_flag = True
    #             except Exception as e:
    #                 msg = f"刪除 {target_folder} 目錄遇到錯誤，錯誤原因： {e}。"
    #                 WSysLog("3", "SyncFolder_RemoveExtraFolder", msg)
        
    #     # 刪除差異檔案
    #     for file in files:
    #         target_file = os.path.join(root, file)
    #         rel_file = os.path.relpath(target_file, target_dir)
    #         source_file_full = os.path.join(source_dir, rel_file)
    #         if not os.path.exists(source_file_full):
    #             try:
    #                 os.remove(target_file)
    #                 msg = f"刪除檔案 {target_file}。"
    #                 WSysLog("1", "SyncFolder_RemoveExtraFile", msg)
    #                 sync_flag = True
    #             except Exception as e:
    #                 msg = f"刪除 {target_file} 檔案時遇到錯誤，錯誤原因： {e}。"
    #                 WSysLog("3", "SyncFolder_RemoveExtraFile", msg)
    if not sync_flag:
        msg = "OneDrive與本地目錄無差異。"
        WSysLog("1", "SyncFolder", msg)

# 本地檔案同步上OneDrive
def UploadOneDrive():
    print()

# OneDrive檔案同步至雲端
def DownloadOneDrive():
    print()

if __name__ == "__main__":
    # GetUserName()
    check_onedrive_path()
    skip_folders = [
        "文件資料(僅存雲端)"
    ]
    sync_folder(os.path.join(Config.OneDrivePath, Config.FolderName), f"{Config.RootDir}/00/", skip_folders)
    # aa = os.path.join(SystemRoot, "654654")
    # print(aa)