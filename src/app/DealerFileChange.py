# -*- coding: utf-8 -*-

'''
檔案說明：經銷商檔案預處理
Writer：Qian
'''
"""
1002322861-康宜
    執行檔名轉換
    執行檔案切割
1002357130-保慶
    執行檔案切割
"""

import os, shutil
import pandas as pd
from datetime import datetime

from Log import WSysLog
from Config import AppConfig

Config = AppConfig()

# 1002322861- 康宜
def change_file_name(dealer_id, folder_path, file_key_word):
    file_names = [file for file in os.listdir(folder_path) \
        if os.path.isfile(os.path.join(folder_path, file))]

    for file in file_names:
        file_path = os.path.join(folder_path, file)
        file_time = os.path.getmtime(file_path)
        file_name, extension = os.path.splitext(file)
        part = file_name.split("_")

        if part[0] == file_key_word:
            if (part[1] == "I") or (part[1] == "S"):
                part[0] = dealer_id
                try:
                    if len(part[2]) == 8:
                        _ = datetime.strptime(part[2], "%Y%m%d")

                    elif len(part[2]) == 12:
                        _ = datetime.strptime(part[2], "%Y%m%d%H%M")
                    else:
                        raise ValueError("日期部分長度不正確。")

                    new_file_name = f"{part[0]}_{part[1]}_{part[2]}{extension}"
                    old_file_path = os.path.join(folder_path, file)
                    new_file_path = os.path.join(folder_path, new_file_name)
                    os.rename(old_file_path, new_file_path)
                    if os.path.exists(new_file_path):
                        msg = f"檔案已變更名稱為：{new_file_name}。"
                        WSysLog("1", "ChangeFileName", msg)
                    os.utime(new_file_path, (file_time, file_time))
                except IndexError:
                    msg = "讀取檔名中的日期時間發生錯誤。錯誤原因:轉換來源不存在。"
                    WSysLog("3", "ChangeFileName", msg)

                except TypeError:
                    msg = "讀取檔名中的日期時間發生錯誤。錯誤原因:轉換來源資料型態錯誤。"
                    WSysLog("3", "ChangeFileName", msg)

                except ValueError as ve:
                    msg = f"讀取檔名中的日期時間發生錯誤。錯誤原因: {str(ve)}"
                    WSysLog("3", "ChangeFileName", msg)

                except FileNotFoundError:
                    msg = f"重新命名錯誤。錯誤原因: 檔案 '{file}' 不存在。"
                    WSysLog("3", "ChangeFileName", msg)

                except PermissionError:
                    msg = "重新命名錯誤。錯誤原因: 權限不足，無法重新命名檔案。"
                    WSysLog("3", "ChangeFileName", msg)

                except Exception as e:
                    msg = f"發生未知錯誤: {e}"
                    WSysLog("3", "ChangeFileName", msg)
            elif part[1] == "INVOICE":
                folder_name = Config.SystemTime.strftime("%Y%m")
                target_path = os.path.join(folder_path, Config.CompleteFolder, folder_name)
                if not os.path.exists(target_path):
                    os.makedirs(target_path)
                    msg = f"已在 {Config.CompleteFolder} 目錄下建立資料夾，資料夾名稱 {folder_name}"
                    WSysLog("1", "ChangeFileName", msg)

                old_file_path = os.path.join(folder_path, file)
                new_file_path = os.path.join(target_path, file)
                shutil.move(old_file_path, new_file_path)
                msg = f"系統移動 {file_name} 至 {Config.CompleteFolder}/{Config.SystemTime.strftime('%Y%m')} 目錄底下。"
                WSysLog("1", "ChangeFileName", msg)

            else:
                # 非約定檔名的處理
                msg = f"{file} 檔案為非約定檔名，系統將刪除該檔案。"
                WSysLog("2", "ChangeFileName", msg)
                os.remove(file_path)
                if not os.path.exists(file_path):
                    msg = f"{file} 檔案已成功刪除。"
                    WSysLog("1", "ChangeFileName", msg)
        else:
            # 非約定檔名的處理
            msg = f"{file} 檔案為非約定檔名，系統將刪除該檔案。"
            WSysLog("2", "ChangeFileName", msg)
            os.remove(file_path)
            if not os.path.exists(file_path):
                msg = f"{file} 檔案已成功刪除。"
                WSysLog("1", "ChangeFileName", msg)

# 將目標欄位的年份變更為西元年
def change_roc_year(folder_path):
    column_keys = [Config.SF_Default_Header[23], Config.SF_Default_Header[26]]
    file_names = [file for file in os.listdir(folder_path) \
        if os.path.isfile(os.path.join(folder_path, file))]
    
    if file_names:
        for file in file_names:
            file_path = os.path.join(folder_path, file)
            file_time = os.path.getmtime(file_path)
            file_name, extension = os.path.splitext(file)
            if extension.lower() == Config.AllowFileExtensions[0]:
                data = pd.read_csv(file_path, dtype = str)

            elif extension.lower() in Config.AllowFileExtensions[1:]:
                data = pd.read_excel(file_path, dtype = str)
                extension = ".xlsx"
                new_file_name = file_name + extension

            else:
                msg = f"{extension} 副檔名不再許可範圍。"
                WSysLog("1", "ChangeRocYear", msg)
                continue

            for column_key in column_keys:
                if column_key in data.columns:
                    for row in range(len(data)):
                        value = str(data.loc[row, column_key])
                        value_part = value.split("/")
                        value_part[0] = str(int(value_part[0]) + 1911)
                        value = "/".join(value_part)
                        data.loc[row, column_key] = value
                else:
                    msg = f"{column_key} 欄位表頭不存在於資料中。"
                    WSysLog("2", "ChangeRocYear", msg)

            if extension.lower() == Config.AllowFileExtensions[0]:
                data.to_csv(file_path, index = False)
                os.utime(file_path, (file_time, file_time))

            elif extension in Config.AllowFileExtensions[1:]:
                os.remove(file_path)
                if not os.path.exists(file_path):
                    msg = f"{file} 檔案已成功刪除。"
                    WSysLog("1", "ChangeRocYear", msg)
                file_path = os.path.join(folder_path, new_file_name)
                data.to_excel(file_path, index=False)
                os.utime(file_path, (file_time, file_time))
        
            msg = f"{file_name} 檔案已過濾出當月資料。"
            WSysLog("1", "ChangeRocYear", msg)
    
    else:
        msg = f"{folder_path} 目標目錄底下無檔案。"
        WSysLog("1", "ChangeRocYear", msg)

# 依據指定欄位值切割檔案內容
def split_file_data(folder_path):
    column_key = Config.SF_Default_Header[23]
    file_names = [file for file in os.listdir(folder_path) \
        if os.path.isfile(os.path.join(folder_path, file))]
    if file_names:
        for file in file_names:
            file_path = os.path.join(folder_path, file)
            file_time = os.path.getmtime(file_path)
            file_name, extension = os.path.splitext(file)
            if extension.lower() == Config.AllowFileExtensions[0]:
                data = pd.read_csv(file_path, dtype = str)

            elif extension.lower() in Config.AllowFileExtensions[1:]:
                data = pd.read_excel(file_path, dtype = str)

            else:
                msg = f"{extension} 副檔名不再許可範圍。"
                WSysLog("1", "SplitFileData", msg)
                continue

            data[column_key] = pd.to_datetime(data[column_key], errors = "coerce")
            part = file_name.split("_")
            data_time = part[2]
            if len(data_time) == 8:
                data_time = datetime.strptime(data_time, "%Y%m%d")
                year_index = data_time.year
                month_index = data_time.month
            elif len(data_time) == 12:
                data_time = datetime.strptime(data_time, "%Y%m%d%H%M")
                year_index = data_time.year
                month_index = data_time.month
            else:
                year_index = Config.Year
                month_index = Config.Month
                msg = "檔名時間格式不對，系統無法處理，將依據系統年月篩選檔案內容。"
                WSysLog("2", "SplitFileData", msg)
            
            # 過濾出對應時間區間的檔案內容
            data = data[(data[column_key].dt.year == year_index) & (data[column_key].dt.month == month_index)]
            # 內容還原成原本的欄位格式YYYY/MM/DD
            data[column_key] = data[column_key].dt.strftime("%Y/%m/%d")

            if extension.lower() == Config.AllowFileExtensions[0]:
                data.to_csv(file_path, index = False, encoding = "UTF-8")
                os.utime(file_path, (file_time, file_time))

            elif extension.lower() in Config.AllowFileExtensions[1:]:
                data.to_excel(file_path, index = False)
                os.utime(file_path, (file_time, file_time))

            msg = f"{file_name} 檔案已過濾出當月資料。"
            WSysLog("1", "SplitFileData", msg)
    else:
        msg = f"{folder_path} 目標目錄底下無檔案。"
        WSysLog("1", "SplitFileData", msg)

# 經銷商檔案預處理主流程
def Preprocessing():
    dealer_list = ["1002322861", "1002317244", "1002357130"]
    folder_path = Config.DealerFolderPath
    for dealer_id in dealer_list:
        # print(f"dealer_id:{dealer_id}")
        path = os.path.join(folder_path, dealer_id)
        # 康宜變更檔名
        if dealer_id == "1002322861":
            change_file_name(dealer_id, path, "NCOSCA")
        elif dealer_id == "1002317244":
            change_file_name(dealer_id, path, "Unimed")
            change_roc_year(path)
        split_file_data(path)
        

if __name__ == "__main__":
    Preprocessing()