# -*- coding: utf-8 -*-

'''
檔案說明：OneDrive上傳或下載檔案
Writer:Qian
'''

"""
需下載目錄 - OneDrive
00_System/Config
00_System/Log
01_BD/01_Masterfile
01_BD/02_Report/底下檔案，子目錄不用
01_BD/03_BA
01_BD/04_DealerInfo
02_Dealer/底下經銷商檔案，00_ChangeFile、00_Completed目錄不用
"""

# 標準庫
import os, shutil

# 自定義函數
from Log import WSysLog
from Config import AppConfig
from SystemConfig import WriteWinUser, WriteOneDrivePath

Config = AppConfig()

# 取得使用者名稱，寫入config
def GetUserName():
    try:
        username = os.getlogin()
        WinUser = Config.WinUser
        if username != WinUser:
            msg = "系統使用者名稱，與紀錄不符，需更新紀錄檔。"
            WSysLog("1", "GetUserName", msg)
            msg = WriteWinUser(username)
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

        try:
            one_drive_path = None
            # 歷遍系統C槽目錄，找尋OneDrive資料夾
            for dirs, folders, _ in os.walk(Config.SystemRoot):
                if Config.OneDeiveFolder in folders:
                    one_drive_path = os.path.join(dirs, Config.OneDeiveFolder)
                    break
            if one_drive_path is None:
                msg = f"未能找到 {Config.OneDeiveFolder} 資料夾。"
                WSysLog("4", "CheckOneDrivePath", msg)
                return False
            if Config.OneDrivePath != one_drive_path:
                msg = WriteOneDrivePath(one_drive_path)
                WSysLog("1", "CheckOenDrivePath", msg)
            return True
        except FileNotFoundError as e:
            msg = f"無法找到 OneDrive 資料夾。錯誤原因： {e}。"
            WSysLog("3", "CheckOenDrivePath", msg)
            return False
        except Exception as e:
            msg = f"搜尋 OneDrive 資料夾時發生意外錯誤。錯誤原因： {e}。"
            WSysLog("4", "CheckOenDrivePath", msg)
            return False
    else:
        msg = "確認OneDrive目錄存在。"
        WSysLog("1", "CheckOneDrivePath", msg)
        return True

# 取得 ba 資料夾名稱
def get_ba_folder_name():
    user = Config.UserConfig
    ba_folders = []
    for user_id in user:
        if user[user_id]["Group"] == "BD_BA":
            ba_id_part = user[user_id]["Description"].split(" ")
            ba_id = ba_id_part[1]
            ba_name = user[user_id]["Name"]
            ba_folders.append(ba_id[:2] + "_" + ba_id[2:] + "_" + ba_name)
    return ba_folders

# 從OneDrive雲端抓取目錄結構至本地
# source_dir是雲端，target_dir是本地
def DownloadOneDrive(source_dir, target_dir, skip_list = []):
    folder_flag, file_flag, dealer_list = False, False, []
    ba_folders = get_ba_folder_name()
    download_folder = [Config.ConfigFolder, Config.MasterFileFolder, Config.ReportFolder, Config.DealerInfoFolder] + ba_folders + Config.DealerList
    not_copy_folder = [Config.ChangeFolder, Config.ErrorReportFolder]
    result = check_onedrive_path()
    if result:
        # 歷遍雲端目錄
        for root, dirs, _ in os.walk(source_dir):
            # 從目錄清單中刪除要跳過的目錄
            dirs[:] = [d for d in dirs if d not in skip_list]

            # 本地同步雲端目錄
            for folder in dirs:
                source_folder = os.path.join(root, folder)
                target_folder = os.path.join(target_dir, os.path.relpath(source_folder, source_dir))
                
                try:
                    if not os.path.exists(target_folder):
                        os.makedirs(target_folder)
                        msg = f"本地建立目錄 {target_folder} 。"
                        WSysLog("1", "DownloadOneDrive_SyncFolders", msg)
                        folder_flag = True
                except OSError as e:
                    msg = f"創建目錄錯誤 {target_folder} : {e} 。"
                    WSysLog("3", "DownloadOneDrive_SyncFolders", msg)

                part = source_folder.split("\\")
                if (not_copy_folder[0] not in part) and (not_copy_folder[1] not in part):
                    # 拷貝特定目錄底下的檔案
                    if folder in download_folder:
                        folder_path = os.path.join(root, folder)
                        file_names = [file for file in os.listdir(folder_path)\
                                    if os.path.isfile(os.path.join(folder_path, file))]

                        for file in file_names:
                            source_file = os.path.join(folder_path, file)
                            target_file = os.path.join(target_dir, os.path.relpath(source_file, source_dir))
                            try:
                                # 本地檔案不存在 或是 本地檔案日期小於雲端日期，拷貝至本地
                                if not os.path.exists(target_file) or os.stat(source_file).st_mtime > os.stat(target_file).st_mtime:
                                    shutil.copy2(source_file, target_file)
                                    if Config.DealerFolder in part:
                                        folder_part = folder_path.split("\\")
                                        dealer_list.append(folder_part[-1])
                                    msg = f"拷貝檔案 {source_file} 至 {target_file}。"
                                    WSysLog("1", "DownloadOneDrive_SyncFiles", msg)
                                    file_flag = True

                            except (IOError, shutil.Error) as e:
                                msg = f"拷貝 {source_file} 至 {target_file} 錯誤，錯誤原因： {e}。"
                                WSysLog("3", "DownloadOneDrive_SyncFiles", msg)

        if not folder_flag:
            msg = "OneDrive 與 本地 目錄結構無差異。"
            WSysLog("1", "DownloadOneDrive", msg)
        if not file_flag:
            msg = "本地檔案已更新至最新。"
            WSysLog("1", "DownloadOneDrive", msg)
        
        return True, list(set(dealer_list))
    else:
        return False, None

# 從本地上傳至 OneDrive 雲端
# source_dir是本地，target_dir是雲端
def UploadOneDrive(source_dir, target_dir):
    not_copy_folder = Config.BAFolder
    result = check_onedrive_path()
    if result:
        folder_flag, file_flag = False, False
        for root, dirs, files in os.walk(source_dir):
            for folder in dirs:
                source_folder = os.path.join(root, folder)
                target_folder = os.path.join(target_dir, os.path.relpath(source_folder, source_dir))
                try:
                    if not os.path.exists(target_folder):
                        os.makedirs(target_folder)
                        msg = f"OneDrive 建立目錄 {target_folder} 。"
                        WSysLog("1", "UploadOneDrive_SyncFolders", msg)
                        folder_flag = True
                except OSError as e:
                    msg = f"創建目錄錯誤 {target_folder} : {e} 。"
                    WSysLog("3", "UploadOneDrive_SyncFolders", msg)
            part = source_folder.split("\\")
            if not_copy_folder not in part:
                for file in files:
                    source_file = os.path.join(root, file)
                    target_file = os.path.join(target_dir,\
                        os.path.relpath(source_file, source_dir))
                    try:
                        if not os.path.exists(target_file) or os.stat(source_file).st_mtime > os.stat(target_file).st_mtime:
                            shutil.copy2(source_file, target_file)
                            msg = f"拷貝 {source_file} 至 {target_file} 。"
                            WSysLog("1","UploadOneDrive_SyncFiles", msg)
                            file_flag = True
                    except (IOError, shutil.Error) as e:
                        msg = f"拷貝 {source_file} 至 {target_file} 錯誤，錯誤原因： {e}。"
                        WSysLog("3", "UploadOneDrive_SyncFiles", msg)
        try:
            for dealer_id in Config.DealerList:
                could_dealer_folder_path = os.path.join(Config.OneDrivePath, Config.FolderName, Config.DealerFolder, dealer_id)
                file_names = [file for file in os.listdir(could_dealer_folder_path)\
                    if os.path.isfile(os.path.join(could_dealer_folder_path, file))]
                for file_name in file_names:
                    file_path = os.path.join(could_dealer_folder_path, file_name)
                    os.remove(file_path)
                    msg = f"系統已移除 {dealer_id} 雲端目錄的 {file_name} 檔案。"
                    WSysLog("1", "UploadOneDrive_ClearCouldDealerFolder", msg)
                msg = f"系統已清空 {dealer_id} 雲端上傳檔案之目錄。"
                WSysLog("1", "UploadOneDrive_ClearCouldDealerFolder", msg)
        except OSError as e:
            msg = f"系統清空經銷商雲端上傳目錄發生OS錯誤。錯誤原因： {e}。"
            WSysLog("3", "UploadOneDrive_ClearCouldDealerFolder", msg)
        except Exception as e:
            msg = f"系統清空經銷商雲端上傳目錄發生未知錯誤。錯誤原因： {e}。"
            WSysLog("3", "UploadOneDrive_ClearCouldDealerFolder", msg)
            
        if not folder_flag:
            msg = "本地目錄與OneDrive目錄已同步到最新。"
            WSysLog("1", "UploadOneDrive", msg)
        
        if not file_flag:
            msg = "雲端檔案已更新至最新。"
            WSysLog("1", "UploadOneDrive", msg)
        return True
    else:
        return False

# 清空本地端檔案
def ClearLocal(source_dir):
    try:
        for root, _, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    if not os.path.exists(file_path):
                        msg = f"系統已從本地 {root} 目錄移除 {file} 檔案。"
                        WSysLog("1", "ClearLocal", msg)
    except OSError as e:
        msg = f"系統清空本地目錄發生錯誤。錯誤原因： {e}。"
        WSysLog("3", "ClearLocal", msg)
    except Exception as e:
        msg = f"系統清空本地目錄發生未知錯誤。錯誤原因： {e}。"
        WSysLog("3", "ClearLocal", msg)

if __name__ == "__main__":
    skip = Config.NotCopyFolder
    could = os.path.join(Config.OneDrivePath, Config.FolderName)
    local = os.path.join(Config.RootDir, Config.FolderName)
    # GetUserName()
    # check_onedrive_path()
    DownloadOneDrive(could, local, skip)
    # UploadOneDrive(local, could)
    # ClearLocal(local)