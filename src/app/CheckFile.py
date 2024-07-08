# -*- coding: utf-8 -*-

'''
檔案說明：確認檔案繳交時間，
檢查檔案副檔名、表頭格式及內容
Writer:Qian
'''

import os, re
import pandas as pd
from itertools import groupby
from datetime import datetime
from operator import itemgetter
from Log import WRecLog, WCheLog
from RecordTable import WriteRawData
from SystemConfig import Config, CheckRule, DealerConf

GlobalConfig = Config()
CheckConfig = CheckRule()
DealerDonfig = DealerConf()

DealerList = DealerDonfig["DealerList"]
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

# 自動切換 panda csv 及 excel 讀取器
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
def determine_file_type(file_path, file_name):
    _, extension = os.path.splitext(file_name)
    sale_file_header_key = set(SF_HeaderKey)
    inventory_file_header_key = set(IF_HeaderKey)
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

# 紀錄檔案繳交狀況，error
def RecordSubmission():
    not_submission, record_dic = [], {}
    for i in DealerList:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        note = {}
        if not file_names:
            not_submission.append(i)
            note = {"Sale":"0/未繳交/無檢查","Inventory":"0/未繳交/無檢查"}
        record_dic[i] = note
        note = {}
        for file_name in file_names:
            file_path = os.path.join(path, file_name)
            file_update_time = os.path.getmtime(file_path)
            formatted_time = datetime.fromtimestamp(file_update_time).strftime('%Y-%m-%d %H:%M:%S')
            file_type = determine_file_type(path, file_name)
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
            for i in error_list:
                f.write(i + "\n")
    
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

    if cell_value:
        key = []
        cell_result = merge_ranges(cell_value)
        for i in cell_result:
            key.append(i)
        error_list.append\
        (f"{'、'.join(cell_result[key[0]])} 及 {'、'.join(cell_result[key[1]])} 內容為空")

    error_message = CheckDataTime(dealer_id, file_dir, file_name)
    error_list += error_message

    if not error_list:
        msg = "檔案內容正確"
        WCheLog("1", "CheckContent", dealer_id, file_name ,msg)
        return msg
    else:
        file = os.path.splitext(file_name)[0]
        file_path = os.path.join(file_dir, f"{file}_content_error.txt")
        with open(file_path, "w", encoding = "UTF-8") as f:
            for i in error_list:
                f.write(i + "\n")
        note = "檔案內容錯誤"
        return note

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
        return error_list
    else:
        for i in invalid_indices:
            cell = data_change_to_excel(df, "Creation Date", i)
            cell_value.append(cell)
            msg = f"{cell} 與檔案更新時間不符合"
            WCheLog("2", "CheckDataTime", dealer_id, file_name, msg)
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])
        return error_list

# 檢查檔案
def CheckFile(have_file_list, record_dic):
    accepted_file_list = []
    for i in have_file_list:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        
        for j in file_names:
            _, j_extension = os.path.splitext(j)
            j_extension = j_extension.lower()
            if j_extension in [".csv", ".xls", ".xlsx"]:
                accepted_file_list.append(j)

        for file in accepted_file_list:
            file_type = determine_file_type(path, file)
            if file_type is not None:
                result, num = CheckHeader(i, path, file, file_type)
                if result:
                    note = CheckContent(i, path, file, file_type)
                    msg = record_dic[i][file_type]
                else:
                    note = "檔案表頭錯誤"
                    msg = record_dic[i][file_type]
            message = f"{num}/{msg}/{note}"
            record_dic[i][file_type] = message
    return record_dic

# 檢查檔案主程式
def RecordAndCheck():
    RecordData = []
    file_list, note = RecordSubmission()
    output_data = CheckFile(file_list, note)
    for i in DealerList:
        note = output_data[i]["Sale"]
        RecordData.append(note)
        note = output_data[i]["Inventory"]
        RecordData.append(note)
    WriteRawData(RecordData)

if __name__ == "__main__":
    RecordAndCheck()