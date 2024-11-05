# -*- coding: utf-8 -*-

'''
檔案說明：EFT雲端檔案透過FTP傳輸
Writer：Qian
'''

# 標準庫
import os
import ftplib

# 自定義函數
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

# ===============================================================================================
'''
"""
改寫EFTFile檔案
"""
# -*- coding: utf-8 -*-

# 標準庫
import os
import ftplib
from functools import wraps

# 自定義函數
from Log import WSysLog
from Mail import SendMail
from Config import AppConfig

Config = AppConfig()

# 程式運作的對應開關。設置Ture則為開啟，False則為關閉
# 此開關用於顯示轉換過程中的相關運作進度
ShowScheduleSwitch = True

# 顯示函數運作進度
def schedule(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            if ShowScheduleSwitch:
                print(f"Function {func.__name__} start.")

            result = func(self, *args, **kwargs)

            if ShowScheduleSwitch:
                print(f"Function {func.__name__} end.")

            return result

        except ftplib.all_errors as e:
            msg = f"{func.__name__} 執行時發生錯誤。錯誤原因：{str(e)}"
            self.log_error(func.__name__, msg)
            raise SystemError(msg) from e

    return wrapper

class eftFunction:
    """
    EFT FTP server相關基礎操作
    """

    # class 區域內變數
    def __init__(self):
        self.retry_count = 0
        self.max_retries = Config.MaxTryRange
        self.remote_path = None
        self.ftp_server = None

        self.login_EFT()
        self.content_test()

    # Info log
    def log_info(self, func_name, message):
        WSysLog("1", func_name, message)

    # Warning log
    def log_warning(self, func_name, message):
        WSysLog("2", func_name, message)

    # Error log
    def log_error(self, func_name, message):
        WSysLog("3", func_name, message)

    # 登入EFT
    @schedule
    def login_EFT(self):
        hostname = "eft.carefusion.com"
        username = "bdservice@coign.com.tw"
        print(f"username:{username}")
        # password = r'\vxGtelcz4'
        password = "C@ign27576941"
        print(f"password:{password}")

        self.ftp_server = ftplib.FTP(hostname)
        self.ftp_server.login(user = username, passwd = password)

        msg = "成功登入EFT雲端。"
        self.log_info("login_EFT", msg)

    # 連線測試
    @schedule
    def content_test(self):
        if not self.ftp_server:
            raise SystemError("尚未登入EFT Server。")

        self.ftp_server.retrlines("LIST")

        msg = "EFT測試連線成功。"
        self.log_info("content_test", msg)

    # 離開EFT
    @schedule
    def quitEFT(self):
        if self.ftp_server:
            self.ftp_server.quit()
            msg = "EFT FTP雲端連線已關閉。"
            self.log_info("quitEFT", msg)

    # 刪除在EFT上的檔案
    @schedule
    def deleteFile(self, remote_file_path):
        if not self.ftp_server:
            raise SystemError("尚未登入EFT Server。")

        self.ftp_server.delete(remote_file_path)
        msg = "成功刪除EFT雲端檔案。"
        self.log_info("deleteFile", f"{msg}FilePath：{remote_file_path}")

    # 切換目錄
    @schedule
    def changePath(self, remote_path):
        if not self.ftp_server:
            raise SystemError("尚未登入EFT Server。")

        self.remote_path = remote_path
        self.ftp_server.cwd(remote_path)
        msg = "成功切換目錄。"
        self.log_info("changePath", msg)

    # 上傳單一檔案
    def upload_single_file(self, file_path, file_name):
        with open(file_path, "rb") as file:
            self.ftp_server.storbinary(f"STOR {file_name}", file)

        # 比對檔案大小，判斷檔案是否上傳成功
        remote_size = self.ftp_server.size(file_name)
        local_size = os.path.getsize(file_path)

        if remote_size != local_size:
            raise ValueError(f"檔案 {file_name} 上傳失敗，大小不一致：本地大小 {local_size}，遠端大小 {remote_size}。")

        return True

    # 處理上傳錯誤
    def handle_upload_errors(self, folder_path, error_files):
        if self.retry_count < self.max_retries:
            msg = f"上傳 {error_files} 至 EFT 雲端發生錯誤，系統嘗試重新上傳(嘗試次數:{self.retry_count + 1})。"
            self.log_error("uploadFile", msg)
            self.retry_count += 1
            self.uploadFile(folder_path, error_files)

        else:
            mail_data = {"FileName": "、".join(error_files)}

            send_info = {
                "Mode": "EFTUploadFileError",
                "DealerID": None,
                "MailData": mail_data,
                "FilesPath": None
            }

            SendMail(send_info)
            msg = f"上傳 {error_files} 至 EFT 雲端失敗，已達最大重試次數。"
            self.log_error("uploadFile", msg)

    # 上傳檔案
    @schedule
    def uploadFile(self, folder_path, file_names):
        if not self.ftp_server:
            raise SystemError("尚未登入EFT Server。")

        if not os.path.exists(folder_path) or not file_names:
            raise ValueError("提供的文件夾路徑不存在或文件名列表為空。")

        upload_flag, error_files = True, []

        for file_name in file_names:
            file_path = os.path.join(folder_path, file_name)

            try:
                upload_flag &= self.upload_single_file(file_path, file_name)

            except Exception as e:
                upload_flag = False
                error_files.append(file_name)
                self.log_error("uploadFile", f"上傳檔案 {file_name} 遇到錯誤：{e}")

        if upload_flag:
            msg = f"成功上傳 {len(file_names)} 份檔案至 EFT雲端：{self.remote_path}。"
            self.log_info("uploadFile", msg)

        else:
            self.handle_upload_errors(folder_path, error_files)

class EFTOperation(eftFunction):
    """
    EFT FTP server整體運作流程
    """

    # class 區域內變數
    def __init__(self):
        super().__init__()

        # 取得目錄底下的檔案
        folder_path = Config.ChangeFolderPath
        self.file_names = self.get_files(folder_path)
        if not isinstance(self.file_names, list):
            raise FileNotFoundError("目標目錄錯誤。")
        self.inventory_file = []
        self.sale_files = []
        self.sort_files()

        self.remote_twm_path = "TaiwanIMS (RW)"
        self.remote_sale_path = "InMarketSales"
        self.remote_inventory_path = "Inventory"

    # 取得目標目錄底下的檔案
    def get_files(self, folder_path):
        # 確保目錄存在
        if not os.path.exists(folder_path):
            msg = f"{folder_path} 目錄不存在。"
            self.log_error("get_files", msg)

            try:
                os.makedirs(folder_path)
                msg = f"已建立 {folder_path} 目錄。"
                self.log_error("get_files", msg)
                return None

            except Exception as e:
                msg = f"建立 {folder_path} 目錄失敗，原因{e}。"
                self.log_error("get_files", msg)
                return False

        # 取得 ChangeFile 目錄底下的檔案
        file_names = [file for file in os.listdir(folder_path) \
            if os.path.isfile(os.path.join(folder_path, file))]

        if len(file_names) == 0:
            msg = f"{folder_path} 來源目錄底下檔案數量為 {len(file_names)}。"
            self.log_info("get_files", msg)

        return file_names

    # 分類檔案
    def sort_files(self):
        if isinstance(self.file_names, list) and not self.file_names:
            for file_name in self.file_names:
                part = file_name.split("_")

                if part[0] == "TWN":
                    self.inventory_file.append(file_name)

                elif part[0] == "IMS":
                    self.sale_files.append(file_name)

        else:
            msg = "指定目錄下無檔案可上傳。"
            self.log_info("sort_files", msg)

    # 上傳銷售檔案
    def uploadSaleFile(self):
        # self.changePath(self.remote_twm_path)
        # self.changePath(self.remote_sale_path)
        # self.content_test()
        if self.sale_files:
            self.changePath(self.remote_twm_path)
            self.changePath(self.remote_sale_path)
            self.uploadFile(".", self.sale_files)

        else:
            return

    # 上傳庫存檔案
    def uploadInventoryFile(self):
        if self.inventory_file:
            print()
        else:
            return
        


# 測試上傳檔案至EFT
def upload_test_file():
    test = eftFunction()
    test.changePath(Config.EFTDir)
    folder_path = "."
    file_names = ["README.md"]
    test.uploadFile(folder_path, file_names)

# 刪除測試上傳EFT的檔案
def delete_test_file():
    test = eftFunction()
    test.changePath(Config.EFTDir)
    test.deleteFile("./README.md")


if __name__ == "__main__":
    # delete_test_file()
    test = EFTOperation()
    test.uploadSaleFile()
    # aa = test.get_files(Config.ChangeFolderPath)
    # print("\\vxGtelcz4")

'''