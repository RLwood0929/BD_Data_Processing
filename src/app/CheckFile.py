# -*- coding: utf-8 -*-

'''
檔案說明：確認檔案繳交時間，
檢查檔案副檔名、表頭格式及內容
Writer:Qian
'''

import os, re
import shutil
import numpy as np
import pandas as pd
from itertools import groupby
from datetime import datetime
from operator import itemgetter
from Log import WRecLog, WCheLog, WSysLog
from SystemConfig import Config, CheckRule, DealerConf
from RecordTable import WriteRawData, WriteSummaryData, WriteNotSubmission

GlobalConfig = Config()
CheckConfig = CheckRule()
DealerDonfig = DealerConf()

DealerList = DealerDonfig["DealerList"]
CompletedDir = GlobalConfig["Default"]["CompletedDir"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]

# 銷售檔案參數
SF_MustHave = CheckConfig["SaleFile"]["MustHave"]
SF_2Choose1 = CheckConfig["SaleFile"]["2Choose1"]
SF_KeyWord = CheckConfig["SaleFile"]["FileKey"]
SF_HeaderKey = CheckConfig["SaleFile"]["HeaderKey"]
# 庫存檔案參數
IF_MustHave = CheckConfig["InventoryFile"]["MustHave"]
IF_2Choose1 = CheckConfig["InventoryFile"]["2Choose1"]
IF_KeyWord = CheckConfig["InventoryFile"]["FileKey"]
IF_HeaderKey = CheckConfig["InventoryFile"]["HeaderKey"]

TargetPath = os.path.join(RootDir, FolderName)
DealerPath = os.path.join(TargetPath, DealerDir)

# 將 index 轉換為 excel 欄
def index_to_excel(index):
    return index + 2

# 將 column 轉換為 excel 列
def column_to_excel(col_index):
    """Convert a zero-based column index to Excel-style letter."""
    result = ""
    while col_index >= 0:
        result = chr(col_index % 26 + 65) + result
        col_index = col_index // 26 - 1
    return result

# 將 data 轉變為 excel 欄位
def data_change_to_excel(data, column, row):
    excel_row = index_to_excel(row)
    excel_column_header = column_to_excel(data.columns.get_loc(column))
    excel_cell_header = f"{excel_column_header}{excel_row}"
    return excel_cell_header

# 調整 error list 顯示結果
def merge_ranges(values):
    grouped_result = {}
    # 透過正則表達式將字母與數字拆除
    parsed_values = [(re.match(r"([A-Z]+)(\d+)", v).groups()) for v in values]
    # 依照字母 A~Z，數字 小 ~ 大 排序
    parsed_values.sort(key = lambda x: (len(x[0]), x[0], int(x[1])))
    grouped_values = {k: [int(num) for _, num in g] \
                      for k, g in groupby(parsed_values, key = itemgetter(0))}
    
    for letter, nums in grouped_values.items():
        nums.sort()
        start = nums[0]
        result = []
        for i in range(1, len(nums)):
            if nums[i] != nums[i - 1] +1:
                result.append(f"{letter}{start}~{nums[i - 1]}" \
                              if start != nums[i - 1] else f"{letter}{start}")
                start = nums[i]
        result.append(f"{letter}{start}~{nums[-1]}" if start != nums[-1] else f"{letter}{start}")
        grouped_result[letter] = result
    return grouped_result

# 自動切換 panda 之 csv 或 excel 讀取器
def read_data(file_path):
    file = os.path.basename(file_path)
    _, file_extension = os.path.splitext(file)
    file_extension = file_extension.lower()
    if file_extension == ".csv":
        df = pd.read_csv(file_path)
        return df
    elif file_extension in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path)
        return df

# 依據檔案名稱選擇 file type
def determine_file_type(dealer_id, file_name):
    for i in range(len(DealerList)):
        if dealer_id == DealerList[i]:
            indx = i + 1
            break
    dsf_header_key = DealerDonfig[f"Dealer{indx}"]["SaleFile"]["KeyWord"]
    dif_header_key = DealerDonfig[f"Dealer{indx}"]["InventoryFile"]["KeyWord"]
    if np.isnan(dsf_header_key):
        sf_header_key = SF_HeaderKey
    else:
        sf_header_key = dsf_header_key
    if np.isnan(dif_header_key):
        if_header_key = IF_HeaderKey
    else:
        sf_header_key = dif_header_key

    file_path = os.path.join(DealerPath, dealer_id)
    _, extension = os.path.splitext(file_name)
    sale_file_header_key = set(sf_header_key)
    inventory_file_header_key = set(if_header_key)
    path = os.path.join(file_path, file_name)
    if extension in [".csv", ".xls", ".xlsx"]:
        data = read_data(path)
        header = data.columns.tolist()
        file_header = set(header)
        sale_result = sale_file_header_key.issubset(file_header)
        inventory_result = inventory_file_header_key.issubset(file_header)
        result = sale_result ^ inventory_result
        parts = file_name.split("_")
        for i in parts:
            if i == SF_KeyWord or i == IF_KeyWord:
                if i == SF_KeyWord and result:
                    return "Sale"
                elif i == IF_KeyWord and result:
                    return "Inventory"
                else:
                    return None
    else:
        return None

# 將檢查錯誤的檔案搬移到 Completed/系統日期 之資料夾底下
def move_check_error_file(dealer_id, file_names):
    today = datetime.today()
    folder_name = today.strftime("%Y%m%d")
    source_path = os.path.join(DealerPath, dealer_id)
    target_path = os.path.join(source_path, CompletedDir, folder_name)
    if not os.path.exists(target_path):
        os.makedirs(target_path)
        msg = f"已在 {CompletedDir} 目錄下建立資料夾，資料夾名稱 {folder_name}"
        WSysLog("1", "MoveCheckErrorFile", msg)

    for file_name in file_names:
        file_source = os.path.join(source_path, file_name)
        file_target = os.path.join(target_path, file_name)
        shutil.move(file_source, file_target)
        if os.path.exists(file_target):
            msg = f"檔案搬移至 {target_path} 成功"
            WSysLog("1", "MoveCheckErrorFile", msg)
        else:
            msg = f"檔案搬移至 {target_path} 失敗"
            WSysLog("2", "MoveCheckErrorFile", msg)

# 紀錄檔案繳交狀況
def RecordSubmission():
    not_submission, record_dic = [], {}
    for i in DealerList:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]

        note = {}
        if not file_names:
            not_submission.append(i)
            record_dic[i] = note

        note = {}
        for file_name in file_names:
            file_path = os.path.join(path, file_name)
            file_update_time = os.path.getmtime(file_path)
            formatted_time = datetime.fromtimestamp(file_update_time).strftime('%Y-%m-%d %H:%M:%S')
            file_type = determine_file_type(i, file_name)

            if file_type is not None:
                msg = f"{file_type} 檔案準時繳交，繳交時間 {formatted_time}"
                WRecLog("1", "RecordSubmission", i, file_name, msg)
                sale_note, inventory_note = {}, {}
                if file_type == "Sale":
                    sale_note = {"Sale":"有繳交"}
                elif file_type == "Inventory":
                    inventory_note = {"Inventory":"有繳交"}
                note.update(sale_note)
                note.update(inventory_note)
            record_dic[i] = note

        if "Sale" not in record_dic[i]:
            record_dic[i]["Sale"] = "0/未繳交/無檢查/0"
        if "Inventory" not in record_dic[i]:
            record_dic[i]["Inventory"] = "0/未繳交/無檢查/0"

    have_submission = list(set(DealerList) - set(not_submission))

    for i in not_submission:
        msg = "檔案未繳交"
        WRecLog("2", "RecordSubmission", i, None, msg)

    return have_submission, record_dic

# 檢查檔案表頭，銷售檔案套用 file_type = Sale；庫存檔案套用 file_type = Inventory。
def CheckHeader(dealer_id, file_dir, file_name, file_type):
    if file_type == "Sale":
        file_must_have = set(SF_MustHave)
        file_2_choose_1 = set(SF_2Choose1)
    elif file_type == "Inventory":
        file_must_have = set(IF_MustHave)
        file_2_choose_1 = set(IF_2Choose1)
    
    flag = False
    file_path = os.path.join(file_dir, file_name)
    data = read_data(file_path)
    max_row = data.shape[0]
    move_list = []
    error_list = []
    header = data.columns.tolist()
    file_header = set(header)
    result1 = file_must_have.issubset(file_header)
    result2 = file_2_choose_1.issubset(file_header)

    if result1 and result2:
        msg = "必要表頭都存在"
        WCheLog("1", "CheckHeader", dealer_id, file_name, msg)
        flag = True
    else:
        header_less = list(file_must_have - file_header)
        msg = f"必要表頭不存在，缺少表頭 {header_less}"
        error_list.append(msg)
        WCheLog("2", "CheckHeader", dealer_id, file_name, msg)
    
    if len(error_list) > 0:
        file = os.path.splitext(file_name)[0]
        file_path = os.path.join(file_dir, f"{file}_header_error.txt")
        with open(file_path, "w", encoding = "UTF-8") as f:
            for i in range(len(error_list)):
                f.write(f"{i+1}. {error_list[i]}\n")
        move_list.append(file_name)
        move_list.append(f"{file}_header_error.txt")
        move_check_error_file(dealer_id, move_list)

    if flag:
        return True, max_row
    else:
        return False, max_row

# 檢查檔案內容，銷售檔案套用 file_type = Sale；庫存檔案套用 file_type = Inventory。
def CheckContent(dealer_id, file_dir, file_name, file_type):
    if file_type == "Sale":
        MustHave = SF_MustHave
        TwoChooseOne = SF_2Choose1
    elif file_type == "Inventory":
        MustHave = IF_MustHave
        TwoChooseOne = IF_2Choose1
    
    file_path = os.path.join(file_dir, file_name)
    data = read_data(file_path)
    error_list = []
    must_have_values = {}
    for i in MustHave:
        must_have_values[i] = data.index[data[i].isna()].tolist()
    
    cell_value = []
    for i in must_have_values:
        for j in must_have_values[i]:
            cell = data_change_to_excel(data, i, j)
            cell_value.append(cell)
            msg = f"{cell} 內容為空"
            WCheLog("2", "CheckContent", dealer_id, file_name ,msg)
    error_num = len(cell_value)
    if cell_value:
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])

    header1 = TwoChooseOne[0]
    header2 = TwoChooseOne[1]
    cell_value = []
    # Original Quantity欄位
    for i in data.index[data[header1].isna()].tolist():
        cell = data_change_to_excel(data, header1, i)
        msg = f"{cell} 內容為空"
        WCheLog("1", "CheckContent", dealer_id, file_name ,msg)

        # Quantity欄位
        if pd.isna(data.at[i, header2]):
            cell2 = data_change_to_excel(data, header2, i)
            msg = f"{cell} 及 {cell2} 內容皆為空"
            cell_value.append(cell)
            cell_value.append(cell2)
            WCheLog("2", "CheckContent", dealer_id, file_name ,msg)
        else:
            cell2 = data_change_to_excel(data, header2, i)
            msg = f"{cell2} 內容有值"
            WCheLog("1", "CheckContent", dealer_id, file_name ,msg)

    error_num = error_num + len(cell_value)
    if cell_value:
        key = []
        cell_result = merge_ranges(cell_value)
        for i in cell_result:
            key.append(i)
        error_list.append\
        (f"{'、'.join(cell_result[key[0]])} 及 {'、'.join(cell_result[key[1]])} 內容為空")

    error_message, e_num = CheckDataTime(dealer_id, file_dir, file_name)
    error_list += error_message
    error_num = error_num + e_num
    move_list = []
    if not error_list:
        msg = "檔案內容正確"
        WCheLog("1", "CheckContent", dealer_id, file_name ,msg)
        return msg, error_num
    else:
        file = os.path.splitext(file_name)[0]
        file_path = os.path.join(file_dir, f"{file}_content_error.txt")
        with open(file_path, "w", encoding = "UTF-8") as f:
            for i in range(len(error_list)):
                f.write(f"{i+1}. {error_list[i]}\n")
        move_list.append(file_name)
        move_list.append(f"{file}_content_error.txt")
        move_check_error_file(dealer_id, move_list)
        note = "檔案內容錯誤"
        return note, error_num

# 檢查檔案內容創建時間與檔案更新時間是否符合，檔案內容創建日期 <= 檔案更新日期
def CheckDataTime(dealer_id, file_dir, file_name):
    Falg = False
    cell_value = []
    error_list = []
    file_path = os.path.join(file_dir, file_name)
    file_update_time = os.path.getmtime(file_path)
    file_update_time = datetime.fromtimestamp(file_update_time).date()
    df = read_data(file_path)
    df["Creation Date"] = pd.to_datetime(df['Creation Date'], format='%Y/%m/%d').dt.date
    df["is_valid"] = df['Creation Date'] <= file_update_time
    invalid_indices = df.index[~df["is_valid"]]
    for i in df["is_valid"]:
        if not i:
            Falg = True

    if not Falg:
        msg = "檔案時間符合"
        WCheLog("1", "CheckDataTime", dealer_id, file_name, msg)
        return error_list, 0
    else:
        for i in invalid_indices:
            cell = data_change_to_excel(df, "Creation Date", i)
            cell_value.append(cell)
            msg = f"{cell} 與檔案更新時間不符合"
            WCheLog("2", "CheckDataTime", dealer_id, file_name, msg)
        error_num = len(cell_value)
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])
        return error_list,  error_num

# 檢查檔案
def CheckFile(have_file_list, record_dic):
    for i in have_file_list:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        accepted_file_list = []
        for j in file_names:
            _, j_extension = os.path.splitext(j)
            j_extension = j_extension.lower()
            if j_extension in [".csv", ".xls", ".xlsx"]:
                accepted_file_list.append(j)

        for file in accepted_file_list:
            file_type = determine_file_type(i, file)
            if file_type is not None:
                result, num = CheckHeader(i, path, file, file_type)
                if result:
                    note, error_num = CheckContent(i, path, file, file_type)
                    msg = record_dic[i][file_type]
                else:
                    note = "檔案表頭錯誤"
                    msg = record_dic[i][file_type]
                    error_num = 0
            message = f"{num}/{msg}/{note}/{error_num}"
            record_dic[i][file_type] = message
    return record_dic

# 檢查檔案主程式
def RecordAndCheck():
    RawData, SummaryData = [], []
    HaveFileList, Note = RecordSubmission()
    output_data = CheckFile(HaveFileList, Note)
    
    for i in DealerList:
        for j in range(2):
            # 寫入Summary
            file_type = "Sale" if j == 0 else "Inventory"
            data = ["0"] * 8
            note = output_data[i][file_type]
            part = note.split("/")
            if part[1] == "有繳交":
                data[0] = 1
            if part[2] in ("檔案內容錯誤", "檔案表頭錯誤"):
                data[2] = 1
            data[1] = part[0]
            data[3] = part[3]
            if data != ["0"] * 8:
                SummaryData = [i, file_type] + data
                print(f"SummaryData:{SummaryData}")
                WriteSummaryData(SummaryData)
            
            # 寫入RawData
            note = output_data[i][file_type]
            RawData.append(note)
    WriteRawData(RawData)

    


if __name__ == "__main__":
    RecordAndCheck()