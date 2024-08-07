# -*- coding: utf-8 -*-

'''
檔案說明：MasterFile、DealerInfo、KAList等檔案初始化建立、合併
Writer：Qian
'''

# 三個合併檔案待寫

import os
import pandas as pd
from Log import WSysLog
from Config import AppConfig
from datetime import datetime
from SystemConfig import WriteFileJson, WriteDealerJson
from openpyxl import Workbook, load_workbook

Config = AppConfig()

# 產生對應的 Excel Column 名稱(一次多個)
def get_excel_colmun_name(file_header):
    return [chr(i % 26+ 65) for i in range(len(file_header))]

# 根據表頭搜尋出 Excel 的 Column 名稱(一次一個)
def search_column_name(file_header, col_name):
    for index in range(len(file_header)):
        if file_header[index] == col_name:
            break
    return chr(index % 26 + 65)

# 設定表格樣式
def excel_style(ws, column_width, fixed_columns):
    for i, width in enumerate(column_width):
        ws.column_dimensions[fixed_columns[i]].width = width
    for col in fixed_columns:
        for row in range(1, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = Config.ExcelStyle

# 從使用者資訊中取出BA名單
def get_BA():
    user = Config.UserConfig
    ba_list = []
    for user_id in user:
        if user[user_id]["Group"] == "BD_BA":
            BA_id_part = user[user_id]["Description"].split(" ")
            BA_id = BA_id_part[1]
            BA_name = user[user_id]["Name"]
            BA_Dealer_id = user[user_id]["ResponsibleDealerID"]
            ba_list.append({"BA_ID":BA_id, "BA_Name":BA_name, "BA_Dealer_ID" : BA_Dealer_id})
    return ba_list

# 依照檔案格式建立空白內容之檔案
def make_format_file(folder_path, file_name, sheet_name, file_header, column_width):
    try:
        file_path = os.path.join(folder_path, file_name)
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(title = sheet_name)
        fixed_columns = get_excel_colmun_name(file_header)
        for col, data in zip(fixed_columns, file_header):
            ws[f"{col}1"] = data
        excel_style(ws, column_width, fixed_columns)
        wb.save(file_path)
        msg = f"成功於 {folder_path} 目錄中建立空白 {file_name} 檔案。"
        WSysLog("1", "MakeFormatFile", msg)
    except Exception as e:
        msg = f"於 {folder_path} 建立空白 {file_name} 檔案時發生錯誤： {e}。"
        WSysLog("3", "MakeFormatFile", msg)

# 將BA負責的經銷商ID寫入DealerInfo
def wright_dealer_info_in_file(file_path, sheet_name, file_header, BA_dealer_id):
    try:
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
        position_list = ["業務窗口","IT人員 or 資料管理人員","專案負責人"]
        content_col = ["Position", "Name", "Mail", "Ex"]
        for row in range(2, (len(BA_dealer_id))*3+1, 3):
            index = int((row - 2) / 3)
            col =  search_column_name(file_header, "Dealer ID")
            ws[f"{col}{row}"] = BA_dealer_id[index]
            
            for j in range(3):
                col = search_column_name(file_header, "Position")
                ws[f"{col}{row + j}"] = position_list[j]
                for col_name in content_col:
                    col = search_column_name(file_header, col_name)
                    ws[f"{col}{row + j}"].alignment = Config.ExcelStyle
            
            for col_name in file_header:
                if col_name in content_col:
                    continue
                col = search_column_name(file_header, col_name)
                ws.merge_cells(f"{col}{row}:{col}{row + j}")
                ws[f"{col}{row}"].alignment = Config.ExcelStyle
        wb.save(file_path)
        msg = f"成功寫入經銷商資訊至 {file_path}。"
        WSysLog("1", "WrightDealerInfoInFile", msg)
    except Exception as e:
        msg = f"寫入經銷商資訊發生錯誤： {e}。"
        WSysLog("3", "WrightDealerInfoInFile", msg)

# 檢查BA目錄底下檔案
def CheckBAFolderFiles():
    BA_list = get_BA()
    # 初始建立旗幟
    flag = False
    master_file = Config.MasterFileName
    dealer_info_file = Config.DealerInfoFileName
    kalist_file = Config.KAListFileName
    json_data = Config.FileConfig
    file_time_dic = {}
    for ba in BA_list:
        file_time = {}
        BA_id = ba["BA_ID"]
        BA_name = ba["BA_Name"]
        BA_dealer_id = ba["BA_Dealer_ID"]
        new_ba_id = BA_id[:2] + "_" + BA_id[2:]
        ba_folder_name = new_ba_id + "_" + BA_name
        folder_path = os.path.join(Config.BAFolderPath, ba_folder_name)

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        master_file_path = os.path.join(folder_path, master_file)
        dealer_info_file_path = os.path.join(folder_path, dealer_info_file)

        # 目標目錄找不到 MasterFile 時須建立一份空的 MasterFile
        if not os.path.exists(master_file_path):
            flag = True
            msg = f"{ba_folder_name}目錄下缺少{master_file}檔案，系統將建立空白檔案。"
            WSysLog("2", "CheckBAFolderFiles", msg)
            sheet_name = Config.MasterFileSheetName
            file_header = Config.MasterFileHeader
            column_width = Config.MasterFileColumnWidth
            make_format_file(folder_path, master_file, sheet_name, file_header, column_width)
            file_write_time = datetime.fromtimestamp(os.path.getatime(master_file_path))
            file_write_time = file_write_time.isoformat()
            json_data["FileInfo"][BA_id]["MasterFile"] = file_write_time
            msg = WriteFileJson(json_data)
            WSysLog("1", "CheckBAFolderFiles", msg)
            file_time["MasterFile"] = None
        else:
            file_time["MasterFile"] = datetime.fromtimestamp(os.path.getatime(master_file_path))

        # 目標目錄找不到 DealerInfo 時須建立一份空的 DealerInfo
        if not os.path.exists(dealer_info_file_path):
            flag = True
            msg = f"{ba_folder_name}目錄下缺少{dealer_info_file}檔案，系統將建立空白檔案。"
            WSysLog("2", "CheckBAFolderFiles", msg)
            sheet_name = Config.DealerInfoFileSheetName
            file_header = Config.DealerInfoFileHeader
            file_header = file_header[1:]
            column_width = Config.DealerInfoFileColumnWidth
            column_width = column_width[1:]
            make_format_file(folder_path, dealer_info_file, sheet_name, file_header, column_width)
            wright_dealer_info_in_file(dealer_info_file_path, sheet_name, file_header, BA_dealer_id)
            file_write_time = datetime.fromtimestamp(os.path.getatime(dealer_info_file_path))
            file_write_time = file_write_time.isoformat()
            json_data["FileInfo"][BA_id]["DealerInfo"] = file_write_time
            msg = WriteFileJson(json_data)
            WSysLog("1", "CheckBAFolderFiles", msg)
            file_time["DealerInfo"] = None
        else:
            file_time["DealerInfo"] = datetime.fromtimestamp(os.path.getatime(dealer_info_file_path))

        # 目標目錄找不到 KAList 時須建立一份空的 KAList
        for ka_dealer in Config.KADealerList:
            if ka_dealer in BA_dealer_id:
                ka_list_file_path = os.path.join(folder_path, kalist_file)
                
                # 目標目錄找不到kalist時須建立一份空的kalist
                if not os.path.exists(ka_list_file_path):
                    flag = True
                    msg = f"{ba_folder_name}目錄下缺少{kalist_file}檔案，系統將建立空白檔案。"
                    WSysLog("2", "CheckBAFolderFiles", msg)
                    sheet_name = Config.KAListFileSheetName
                    file_header = Config.KAListFileHeader
                    for i in range(len(Config.DealerList)):
                        if Config.DealerList[i] == ka_dealer:
                            index = i + 1
                            break
                    dealer_name = Config.DealerConfig[f"Dealer{index}"]["DealerName"]
                    sheet_name = ka_dealer + "_" + sheet_name
                    file_header[0] = dealer_name + file_header[0]
                    column_width = Config.KAListFileColumnWidth
                    make_format_file(folder_path, kalist_file, sheet_name, file_header, column_width)
                    file_write_time = datetime.fromtimestamp(os.path.getatime(ka_list_file_path))
                    file_write_time = file_write_time.isoformat()
                    json_data["FileInfo"][BA_id]["KAList"] = file_write_time
                    msg = WriteFileJson(json_data)
                    file_time["KAList"] = None
                else:
                    file_time["KAList"] = datetime.fromtimestamp(os.path.getatime(ka_list_file_path))

        file_time_dic[BA_id] = file_time
    
    if not flag:
        msg = "BA目錄底下應有檔案都存在。" 
        WSysLog("1", "CheckBAFolderFiles", msg)
        return file_time_dic
    else:
        return file_time_dic
    
# 依據BA人數調整files.json格式
def check_ba():
    flag = True
    BA_list = get_BA()
    json_data = Config.FileConfig

    for ba in BA_list:
        BA_id = ba["BA_ID"]
        if BA_id not in json_data["FileInfo"]:
            flag = False
            json_data["FileInfo"][BA_id] = {"MasterFile":None, "DealerInfo":None, "KAList":None}

    if flag:
        msg = "files.json資料無異動。"
        WSysLog("1", "CheckBA", msg)
        return json_data
    else:
        msg = WriteFileJson(json_data)
        WSysLog("1", "CheckBA", msg)
        return json_data

# 系統合併 masterfile 檔案
def merge_masterfile(aa):
    print(aa)

# 系統合併 DealerInfo 檔案 #
def merge_dealerinfo(write_new_list):
    BA_list = get_BA()
    dealer_list = Config.DealerList
    folder_path = Config.DealerInfoPath
    file_name = Config.DealerInfoFileName
    sheet_name = Config.DealerInfoFileSheetName
    file_header = Config.DealerInfoFileHeader
    column_width = Config.DealerInfoFileColumnWidth
    general_data_path = os.path.join(folder_path, file_name)
    content_col = ["Position", "Name", "Mail", "Ex"]
    

    # 總表存放目錄底下沒有總表檔案時，直接合併全部產生一份檔案
    if not os.path.exists(general_data_path):
        msg = f"{folder_path} 目錄底下無 {file_name} 總表檔案，系統將重新產生。"
        WSysLog("2", "MergeDealerInfo", msg)
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(title = sheet_name)
        dataframes, dealers, dealer_part, dealer_part_dct = [], [], [], {}

        for ba in BA_list:
            BA_folder_name = ba["BA_ID"][:2] + "_" + ba["BA_ID"][2:] + "_" + ba["BA_Name"]
            file_path = os.path.join(Config.BAFolderPath, BA_folder_name, file_name)
            dealer_data = pd.read_excel(file_path, dtype = str)
            part = [dealer_data.iloc[i:i+3].reset_index(drop=True) for i in range(0, len(dealer_data), 3)]
            dealer_part.extend(part)
            dealers.extend(ba["BA_Dealer_ID"])
            for dealer in ba["BA_Dealer_ID"]:
                if dealer not in dealer_list:
                    dealer_list.append(dealer)
                    result = WriteDealerJson(dealer_list)
                    if result:
                        msg = f"將 {dealer} 加入 DealerList 紀錄中。"
                        WSysLog("1", "WriteDealerJson", msg)

        for i in range(len(dealer_part)):
            dealer_part_dct[dealer_part[i]["Dealer ID"][0]] = dealer_part[i]

        for dealer in dealer_part_dct:
            dealer_part_dct[dealer].insert(0, "ID", None)
            dealer_part_dct[dealer].loc[0, "ID"] = f"Dealer{(dealer_list.index(dealer_part_dct[dealer]['Dealer ID'][0]) + 1)}"
            dataframes.append(dealer_part_dct[dealer])
        
        conbined_df = pd.concat(dataframes, ignore_index = True)
        fixed_columns = get_excel_colmun_name(file_header)
        for col, data in zip(fixed_columns, file_header):
            ws[f"{col}1"] = data
        excel_style(ws, column_width, fixed_columns)

        # # 將資料寫入工作表
        for row in range(len(conbined_df)):
            for col_name in conbined_df.columns.values:
                col = search_column_name(file_header, col_name)
                ws[f"{col}{row + 2}"] = conbined_df[col_name][row]
                ws[f"{col}{row + 2}"].alignment = Config.ExcelStyle
        
        for row in range(2, len(conbined_df) + 1, 3):
            for col_name in conbined_df.columns.values:
                if col_name not in content_col:
                    col = search_column_name(file_header, col_name)
                    ws.merge_cells(f"{col}{row}:{col}{row + 2}")
        wb.save(general_data_path)
    
    else:
        wb = load_workbook(general_data_path)
        ws = wb[sheet_name]
        dataframes, dealers, dealer_part, dealer_part_dct = [], [], [], {}

        for ba_id in write_new_list:
            for ba in BA_list:
                dealers.extend(ba["BA_Dealer_ID"])
                for dealer in ba["BA_Dealer_ID"]:
                    if dealer not in dealer_list:
                        dealer_list.append(dealer)
                        result = WriteDealerJson(dealer_list)

                        if result:
                            msg = f"將 {dealer} 加入 DealerList 紀錄中。"
                            WSysLog("1", "WriteDealerJson", msg)

                if ba["BA_ID"] == ba_id:
                    BA_folder_name = ba["BA_ID"][:2] + "_" + ba["BA_ID"][2:] + "_" + ba["BA_Name"]
                    break
            file_path = os.path.join(Config.BAFolderPath, BA_folder_name, file_name)
            dealer_data = pd.read_excel(file_path, dtype = str)
            dealer_data = dealer_data.applymap(lambda x: None if pd.isna(x) else x)
            part = [dealer_data.iloc[i:i+3].reset_index(drop=True) for i in range(0, len(dealer_data), 3)]
            dealer_part.extend(part)
            dealers.extend(ba["BA_Dealer_ID"])

        for i in range(len(dealer_part)):
            dealer_part_dct[dealer_part[i]["Dealer ID"][0]] = dealer_part[i]
        
        for dealer in dealer_part_dct:
            dealer_part_dct[dealer].insert(0, "ID", None)
            dealer_part_dct[dealer].loc[0, "ID"] = f"Dealer{(dealer_list.index(dealer_part_dct[dealer]['Dealer ID'][0]) + 1)}"

            index = dealer_list.index(dealer)
            row = ((index - 2) * 3) + 2
            for col_name in dealer_part_dct[dealer].columns.values:
                col = search_column_name(file_header, col_name)
                for i in range(len(dealer_part_dct[dealer])):
                    if ws[f"{col}{row + i}"].value != dealer_part_dct[dealer][col_name][i]:
                        msg = f"更新總表 {col}{row + i} 資訊為： {dealer_part_dct[dealer][col_name][i]}。"
                        WSysLog("1", "MergeDealerInfo", msg)
                        ws[f"{col}{row + i}"] = dealer_part_dct[dealer][col_name][i]
        # 樣式要調整 #
        wb.save(general_data_path)
        


# 系統合併 KAList 檔案
def merge_kalist(aa):
    print()

# 檢查檔案資訊主程式
def CheckFileInfo():
    master_new, kalist_new, dealerinfo_new = [], [], []
    os_file_time = CheckBAFolderFiles()
    json_data = check_ba()
    record_file_time = json_data["FileInfo"]

    # 取得作業系統檔案更新時與紀錄中比對，判定是否需要更新
    for BA_id, files_time in os_file_time.items():
        for file_name, file_time in files_time.items():
            if file_time:
                last_file_time = record_file_time[BA_id][file_name]
                last_file_time = datetime.fromisoformat(last_file_time)

                if last_file_time != file_time:
                    if file_time > last_file_time:
                        msg = f"{BA_id} 的 {file_name} 檔案內容更新，系統需更新 {file_name} 總表。"
                        WSysLog("1", "CheckFileInfo", msg)
                        file_time = file_time.isoformat()
                        record_file_time[BA_id][file_name] = file_time
                        json_data["FileInfo"] = record_file_time
                        msg = WriteFileJson(json_data)
                        WSysLog("1", "CheckFileInfo", msg)

                        if file_name == "MasterFile":
                            master_new.append(BA_id)
                        elif file_name == "DealerInfo":
                            dealerinfo_new.append(BA_id)
                        elif file_name == "KAList":
                            kalist_new.append(BA_id)

                    else:
                        msg = f"{BA_id} 的 {file_name}檔案異動時間異常，小於系統紀錄時間。"
                        WSysLog("3", "CheckFileInfo", msg)

                else:
                    msg = f"系統偵測 {BA_id} 的 {file_name} 檔案無更新。"
                    WSysLog("1", "CheckFileInfo", msg)
            else:
                msg = f"系統初始化建立 {file_name} 檔案於 {BA_id} 中，請填寫完畢後，在執行一次系統。"
                WSysLog("1", "CheckFileInfo", msg)

    if master_new:
        merge_masterfile(master_new)
    if dealerinfo_new:
        merge_dealerinfo(dealerinfo_new)
    if kalist_new:
        merge_kalist(kalist_new)

if __name__ == "__main__":
    # CheckFileInfo()
    aa = ["BA01"]
    merge_dealerinfo(aa)