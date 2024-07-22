# -*- coding: utf-8 -*-

'''
檔案說明：確認檔案繳交時間，
檢查檔案副檔名、表頭格式及內容
Writer:Qian
'''

import os, re
import shutil
import pandas as pd
from itertools import groupby
from operator import itemgetter
from datetime import datetime, timedelta
from Log import WRecLog, WCheLog, WSysLog
from SystemConfig import Config, CheckRule, DealerConf, DealerFormatConf
from RecordTable import WriteRawData, WriteSummaryData, WriteNotSubmission, GetSummaryData

SystemTime = datetime.now()
GlobalConfig = Config()
CheckConfig = CheckRule()
DealerConfig = DealerConf()
DFormatConfig = DealerFormatConf()

DealerList = DealerConfig["DealerList"]
CompletedDir = GlobalConfig["Default"]["CompletedDir"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
MonthlyFileRange = GlobalConfig["App"]["MonthlyFileRange"] if GlobalConfig["App"]["MonthlyFileRange"]\
    else GlobalConfig["Default"]["MonthlyFileRange"]
MonthlyFileRange = int(MonthlyFileRange)

# 銷售檔案參數
SF_MustHave = CheckConfig["SaleFile"]["MustHave"]
SF_2Choose1 = CheckConfig["SaleFile"]["2Choose1"]
SF_Default_Header = DFormatConfig["Defualt"]["SaleFileHeader"]

# 庫存檔案參數
IF_MustHave = CheckConfig["InventoryFile"]["MustHave"]
IF_2Choose1 = CheckConfig["InventoryFile"]["2Choose1"]
IF_Default_Header = DFormatConfig["Defualt"]["InventoryFileHeader"]

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
            index = i + 1
            break

    SF_Header = DealerConfig[f"Dealer{index}"]["SaleFile"]["FileHeader"]
    IF_Header = DealerConfig[f"Dealer{index}"]["InventoryFile"]["FileHeader"]
    
    file_path = os.path.join(DealerPath, dealer_id)
    _, extension = os.path.splitext(file_name)
    sale_file_header = set(SF_Header)
    inventory_file_header = set(IF_Header)
    path = os.path.join(file_path, file_name)
    if extension in [".csv", ".xls", ".xlsx"]:
        data = read_data(path)
        file_header = set(data.columns.tolist())
        sale_result =  sale_file_header == file_header
        inventory_result = inventory_file_header == file_header
        if sale_result != inventory_result:
            if sale_result:
                return "Sale"
            elif inventory_result:
                return "Inventory"
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

# 日繳檔案紀錄
def DailyFile(dealer_id, file_name):
    start_date = SystemTime.date() - timedelta(days = 1)
    end_date = SystemTime.date()
    file_path = os.path.join(DealerPath, dealer_id)
    path = os.path.join(file_path,file_name)
    df = read_data(path)
    flag = False
    # 確認Creation Date欄位存在
    check_col = SF_MustHave[-1]
    check_col = check_col.replace(" ","").lower()
    df.columns = df.columns.str.replace(" ","",regex = False).str.lower()
    if check_col in df.columns:
        df[check_col] = pd.to_datetime(df[check_col], format='%Y/%m/%d').dt.date
        df["is_conform"] = (start_date <= df[check_col]) & (df[check_col] <= end_date)
        for i in df["is_conform"]:
            if not i:
                flag = True
        if not flag:
            msg = "檔案內容Creation Date時間符合繳交區間。"
            WRecLog("1", "DailyFile", dealer_id, file_name, msg)
        else:
            not_conform = df[df["is_conform"] == False]
            index = not_conform.index
            for i in index:
                row = data_change_to_excel(df, check_col, i)
                msg = f"檔案內容 {row} 時間不符合繳交區間。"
                WRecLog("2", "DailyFile", dealer_id, file_name, msg)
    else:
        msg = "表頭中無Creation Date欄位，無法檢查。"
        WRecLog("2", "DailyFile", dealer_id, file_name, msg)

# 月繳檔案紀錄
def MonthlyFile(dealer_id, file_name):
    end_date = SystemTime.date().replace(day = 1) - timedelta(days = 1)
    start_date = end_date.replace(day = 1)
    file_path = os.path.join(DealerPath, dealer_id)
    path = os.path.join(file_path, file_name)
    file_type = determine_file_type(dealer_id, file_name)
    df = read_data(path)
    flag = False
    # 確認Creation Date欄位存在
    check_col = SF_MustHave[-1]
    check_col = check_col.replace(" ","").lower()
    df.columns = df.columns.str.replace(" ","",regex = False).str.lower()
    if check_col in df.columns:
        df[check_col] = pd.to_datetime(df[check_col], format='%Y/%m/%d').dt.date
        df["is_conform"] = (start_date <= df[check_col]) & (df[check_col] <= end_date)
        df["Year"] = df[check_col].dt.year
        creation_year = df["Year"][2]
        df["Month"] = df[check_col].dt.month
        creation_month = df["Month"][2]
        date_due = f"{creation_year}-{creation_month + 1}-05"
        for i in df["Year"]:
            if creation_year != i:
                msg = "檔案內容Creation Date非同一年份"
                WRecLog("2", "MonthlyFile", dealer_id, file_name, msg)
                return False, None
        for i in df["Month"]:
            if creation_month != i:
                msg = "檔案內容Creation Date非同一月份"
                WRecLog("2", "MonthlyFile", dealer_id, file_name, msg)
                return False, None
        for i in df["is_conform"]:
            if not i:
                flag = True
        if not flag:
            msg = "檔案內容Creation Date時間符合繳交區間。"
            WRecLog("1", "MonthlyFile", dealer_id, file_name, msg)
            return True, None
        else:
            sub_data = [dealer_id, file_type, file_name, date_due ,"已繳交", "無檢查", None, None]
            WriteNotSubmission(sub_data)
            return "ReSubmission", date_due
    else:
        msg = "表頭中無Creation Date欄位，無法檢查。"
        WRecLog("2", "MonthlyFile", dealer_id, file_name, msg)
        return False, None
    
# 紀錄檔案繳交狀況
def RecordSubmission():
    not_submission, record_dic = [], {}
    for i in DealerList:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        for index in range(len(DealerList)):
            if DealerList[index] == i:
                index += 1
                break

        sale_file_cycle = DealerConfig[f"Dealer{index}"]["SaleFile"]["PaymentCycle"]
        inventory_file_cycle = DealerConfig[f"Dealer{index}"]["InventoryFile"]["PaymentCycle"]

        note = {}
        if not file_names:
            not_submission.append(i)
            record_dic[i] = note

        note, resubmission_list, date_due_list = {}, [], []
        for file_name in file_names:
            file_path = os.path.join(path, file_name)
            file_update_time = os.path.getmtime(file_path)
            formatted_time = datetime.fromtimestamp(file_update_time).strftime('%Y-%m-%d %H:%M:%S')
            file_type = determine_file_type(i, file_name)

            if file_type is not None:
                msg = f"{file_type} 檔案準時繳交，繳交時間 {formatted_time}"
                WRecLog("1", "RecordSubmission", i, file_name, msg)
                sale_note, inventory_note = {}, {}
                # 銷售檔案
                if file_type == "Sale":
                    sale_note = {"Sale":"有繳交/無補繳"}
                    if sale_file_cycle == "D":
                        DailyFile(i, file_name)
                    else:
                        result, date_due = MonthlyFile(i, file_name)
                        if result == "ReSubmission":
                            sale_note = {"Sale":"有繳交/補繳交"}
                            resubmission_list.append(file_name)
                            date_due_list.append(date_due)
                # 庫存檔案
                elif file_type == "Inventory":
                    inventory_note = {"Inventory":"有繳交/無補繳"}
                    if inventory_file_cycle == "D":
                        DailyFile(i, file_name)
                    else:
                        result, date_due = MonthlyFile(i, file_name)
                        if result == "ReSubmission":
                            inventory_note = {"Inventory":"有繳交/補繳交"}
                            resubmission_list.append(file_name)
                            date_due_list.append(date_due)
                note.update(sale_note)
                note.update(inventory_note)
            record_dic[i] = note

        if "Sale" not in record_dic[i]:
            record_dic[i]["Sale"] = "0/0/未繳交/無補繳/無檢查/0"
        if "Inventory" not in record_dic[i]:
            record_dic[i]["Inventory"] = "0/0/未繳交/無補繳/無檢查/0"

    have_submission = list(set(DealerList) - set(not_submission))
    for i in not_submission:
        msg = "檔案未繳交"
        WRecLog("2", "RecordSubmission", i, None, msg)
    return have_submission, record_dic, resubmission_list, date_due_list

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
        msg = "檔案時間符合。"
        WCheLog("1", "CheckDataTime", dealer_id, file_name, msg)
        return error_list, 0
    else:
        for i in invalid_indices:
            cell = data_change_to_excel(df, "Creation Date", i)
            cell_value.append(cell)
            msg = f"{cell} 與檔案更新時間不符合。"
            WCheLog("2", "CheckDataTime", dealer_id, file_name, msg)
        error_num = len(cell_value)
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])
        return error_list,  error_num

# 檢查檔案
def CheckFile(have_file_list, record_dic, re_submission, date_due_list):
    for i in have_file_list:
        path = os.path.join(DealerPath, i)
        file_names = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        accepted_file_list = []
        for j in file_names:
            _, j_extension = os.path.splitext(j)
            j_extension = j_extension.lower()
            if j_extension in [".csv", ".xls", ".xlsx"]:
                accepted_file_list.append(j)
        
        for index in range(len(DealerList)):
            if i == index:
                index += 1
                break
        sale_file_cycle = DealerConfig[f"Dealer{index}"]["SaleFile"]["PaymentCycle"]
        inventory_file_cycle = DealerConfig[f"Dealer{index}"]["InventoryFile"]["PaymentCycle"]

        for file in accepted_file_list:
            re_sub_file = None
            for j in range(len(re_submission)):
                if file == re_submission[j]:
                    re_sub_file = file
                    date_due = date_due_list[j]
                    break
            file_type = determine_file_type(i, file)
            if file_type is not None:
                result, num = CheckHeader(i, path, file, file_type)
                if result:
                    note, error_num = CheckContent(i, path, file, file_type)
                    msg = record_dic[i][file_type]
                    # 月繳檔案
                    if (file_type == "Sale" and sale_file_cycle == "M" and note == "檔案內容錯誤") or \
                    (file_type == "Inventory" and inventory_file_cycle == "M" and note == "檔案內容錯誤") :
                        re_seb_data = [i, file_type, file, SystemTime.date().replace(day = 5), "已繳交", note, None, None]
                        WriteNotSubmission(re_seb_data)
                else:
                    note = "檔案表頭錯誤"
                    msg = record_dic[i][file_type]
                    error_num = 0
                    if (file_type == "Sale" and sale_file_cycle == "M") or \
                    (file_type == "Inventory" and inventory_file_cycle == "M"):
                        re_seb_data = [i, file_type, file, SystemTime.date().replace(day = 5), "已繳交", note, None, None]
                        WriteNotSubmission(re_seb_data)
                if re_sub_file is not None:
                    re_seb_data = [i, file_type, file, date_due, "已補繳", None, SystemTime, note]
                    WriteNotSubmission(re_seb_data)
            message = f"{num}/{msg}/{note}/{error_num}"
            record_dic[i][file_type] = message
    return record_dic

# 檢查檔案主程式
def RecordAndCheck():
    RawData, SummaryData = [], []
    HaveFileList, Note, ReSubmissionList, DateDueList = RecordSubmission()
    output_data = CheckFile(HaveFileList, Note, ReSubmissionList, DateDueList)
    
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
                WriteSummaryData(SummaryData)

            if SystemTime == SystemTime.replace(day = MonthlyFileRange):
                sum_data = GetSummaryData()
                monthly_data = sum_data[sum_data["檔案繳交週期"] == "M"]
                not_sub_file = monthly_data[monthly_data["當月繳交次數"]== 0].reset_index(drop = True)
                not_sub_list = []
                for k in range(len(not_sub_file)):
                    to_list = not_sub_file.loc[k]
                    to_list = list(to_list.to_dict().values())
                    not_sub_list.append(to_list)
                for k in not_sub_list:
                    not_sub_data = [k[0],k[2],None, SystemTime.date().replace(day = MonthlyFileRange), "未繳交", "無檢查", None, None]
                    WriteNotSubmission(not_sub_data)

            # 寫入RawData
            note = output_data[i][file_type]
            RawData.append(note)
    WriteRawData(RawData)

if __name__ == "__main__":
    RecordAndCheck()

"""
---------------------------------------------------------------------------------------------------
"""

import os, re
import shutil
import pandas as pd
from itertools import groupby
from operator import itemgetter
from datetime import datetime, timedelta
from Log import WSysLog, WRecLog, WCheLog
from RecordTable import WriteSubRawData, WriteNotSubmission
from SystemConfig import Config, DealerConf, DealerFormatConf, CheckRule, SubRecordJson

SystemTime = datetime.now()
GlobalConfig = Config()
CheckConfig = CheckRule()
DealerConfig = DealerConf()
DealerFormatConfig = DealerFormatConf()

DealerList = DealerConfig["DealerList"]
CompletedDir = GlobalConfig["Default"]["CompletedDir"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
MonthlyFileRange = GlobalConfig["App"]["MonthlyFileRange"] if GlobalConfig["App"]["MonthlyFileRange"]\
    else GlobalConfig["Default"]["MonthlyFileRange"]
MonthlyFileRange = int(MonthlyFileRange)
AllowFileExtensions = GlobalConfig["Default"]["AllowFileExtensions"]

# 銷售檔案參數
SF_MustHave = CheckConfig["SaleFile"]["MustHave"]
SF_2Choose1 = CheckConfig["SaleFile"]["2Choose1"]
SF_Default_Header = DealerFormatConfig["Defualt"]["SaleFileHeader"]

# 庫存檔案參數
IF_MustHave = CheckConfig["InventoryFile"]["MustHave"]
IF_2Choose1 = CheckConfig["InventoryFile"]["2Choose1"]
IF_Default_Header = DealerFormatConfig["Defualt"]["InventoryFileHeader"]

DealerPath = os.path.join(RootDir, FolderName, DealerDir)

# 自動切換 panda 之 csv 或 excel 讀取器
def read_data(file_path):
    file = os.path.basename(file_path)
    _, file_extension = os.path.splitext(file)
    file_extension = file_extension.lower()
    if file_extension == AllowFileExtensions[0]:
        df = pd.read_csv(file_path)
        return df, df.shape[0]
    elif file_extension in AllowFileExtensions[1:]:
        df = pd.read_excel(file_path)
        return df, df.shape[0]
    
# 決定檔案類型
def decide_file_type(dealer_id, file_name):
    for i in range(len(DealerList)):
        if dealer_id == DealerList[i]:
            index = i + 1
            break
    
    # 抓取與經銷商協定的表頭
    sale_file_header = DealerConfig[f"Dealer{index}"]["SaleFile"]["FileHeader"]
    inventory_file_header = DealerConfig[f"Dealer{index}"]["InventoryFile"]["FileHeader"]

    folder_path = os.path.join(DealerPath, dealer_id)
    _, extension = os.path.splitext(file_name)
    sale_file_header = set(sale_file_header)
    inventory_file_header = set(inventory_file_header)
    file_path = os.path.join(folder_path, file_name)
    if extension in AllowFileExtensions:
        data, max_row = read_data(file_path)
        file_header = set(data.columns.tolist())
        sale_result =  sale_file_header == file_header
        inventory_result = inventory_file_header == file_header
        if sale_result != inventory_result:
            if sale_result:
                return "Sale", max_row
            elif inventory_result:
                return "Inventory", max_row
            else:
                return None, None
    else:
        os.remove(file_path)
        msg = f"檔案 {file_name} 副檔名不符合，系統已刪除該檔案。"
        WRecLog("2", "DecideFileType", dealer_id, file_name, msg)
        return None, None

# 寫入未繳交名單
def write_not_sub_record():
    record_data = SubRecordJson("Read", None)
    for i in range(len(DealerList)):
        index = i + 1
        sale_cycle = DealerConfig[f"Dealer{index}"]["SaleFile"]["PaymentCycle"]
        inventory_cycle = DealerConfig[f"Dealer{index}"]["InventoryFile"]["PaymentCycle"]
        dealer_id = DealerConfig[f"Dealer{index}"]["DealerID"]
        
        # 銷售日繳
        if (sale_cycle == "D") and (not record_data[f"Dealer{index}"]["SaleFile"]):
            file_extencion = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
            write_data = {
                "經銷商ID":dealer_id,
                "檔案類型":"Sale",
                "缺繳(待補繳)檔案名稱":f"{dealer_id}_S_{SystemTime.date()}.{file_extencion}",
                "檔案狀態":"未繳交",
                "應繳時間":SystemTime.date(),
                "檔案檢查結果":"未檢查"
            }
            WriteNotSubmission(write_data)
        
        # 庫存日繳
        if (inventory_cycle == "D") and (not record_data[f"Dealer{index}"]["InventoryFile"]):
            file_extencion = DealerConfig[f"Dealer{index}"]["InventoryFile"]["Extension"]
            write_data = {
                "經銷商ID":dealer_id,
                "檔案類型":"Inventory",
                "缺繳(待補繳)檔案名稱":f"{dealer_id}_I_{SystemTime.date()}.{file_extencion}",
                "檔案狀態":"未繳交",
                "應繳時間":SystemTime.date(),
                "檔案檢查結果":"未檢查"
            }
            WriteNotSubmission(write_data)
        
        # 銷售月繳
        if (sale_cycle == "M") and (SystemTime == (SystemTime.replace(day = MonthlyFileRange))):
            if not record_data[f"Dealer{index}"]["SaleFile"]:
                file_extencion = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
                write_data = {
                    "經銷商ID":dealer_id,
                    "檔案類型":"Sale",
                    "缺繳(待補繳)檔案名稱":f"{dealer_id}_S_{SystemTime.date()}.{file_extencion}",
                    "檔案狀態":"未繳交",
                    "應繳時間":f"{SystemTime.date().replace(day = 1)} ~ {SystemTime.date().replace(day = MonthlyFileRange)}",
                    "檔案檢查結果":"未檢查"
                }
                WriteNotSubmission(write_data)
        
        # 庫存月繳
        if (inventory_cycle == "M") and (SystemTime == (SystemTime.replace(day = MonthlyFileRange))):
            if not record_data[f"Dealer{index}"]["InventoryFile"]:
                file_extencion = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
                write_data = {
                    "經銷商ID":dealer_id,
                    "檔案類型":"Sale",
                    "缺繳(待補繳)檔案名稱":f"{dealer_id}_I_{SystemTime.date()}.{file_extencion}",
                    "檔案狀態":"未繳交",
                    "應繳時間":f"{SystemTime.date().replace(day = 1)} ~ {SystemTime.date().replace(day = MonthlyFileRange)}",
                    "檔案檢查結果":"未檢查"
                }
                WriteNotSubmission(write_data)

# 檢查檔案命名格式 #
def check_file_name_format(dealer_id, file_name, file_extencion):
    flag = True
    file_name_part = re.split(r"[._]" ,file_name)
    if file_name_part[0] not in DealerList:
        flag = False
        msg = "檔名內容錯誤，經銷商ID不再範圍中。"
        WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    if (file_name_part[1] != "S") and (file_name_part[1] != "I"):
        flag = False
        msg = "檔名內容錯誤，檔案類型不再範圍中。"
        WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    file_name_part2 = file_name_part[2]
    if len(file_name_part2) == 8:
        try:
            file_time = datetime.strptime(file_name_part2, "%Y%m%d")
        except ValueError as  e:
            flag = False
            msg = f"檔名內容錯誤，時間內容錯誤 {e}。"
            WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    else:
        try:
            file_time = datetime.strptime(file_name_part2, "%Y%m%d%H%M")
            file_time = file_time.date()
        except ValueError as  e:
            flag = False
            msg = f"檔名內容錯誤，時間內容錯誤 {e}。"
            WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    if file_name_part[-1] != file_extencion:
        flag = False
        msg = "檔案副檔名錯誤。"
        WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    
    if not flag:
        try:
            file_path = os.path.join(DealerPath, dealer_id, file_name)
            #os.remove(file_path)
            msg = f"已移除檔案 {file_name}。"
            WSysLog("1", "RecordDealerfiles", msg)
        except Exception as e:
            msg = f"移除檔案 {file_name} 時發生未知錯誤： {e}。"
            WSysLog("2", "RecordDealerfiles", msg)
    return file_time

# 紀錄檔案繳交主程式
def RecordDealerFiles():
    not_submission, data_dic,  = [], {}
    for dealer_id in DealerList:
        # 抓取經銷商目錄底下檔案
        dealer_path = os.path.join(DealerPath, dealer_id)
        file_names = [file for file in os.listdir(dealer_path) \
                      if os.path.isfile(os.path.join(dealer_path, file))]
        
        for i in range(len(DealerList)):
            if DealerList[i] == dealer_id:
                index = i + 1
                break

        # 取得經銷商檔案繳交週期
        sale_cycle = DealerConfig[f"Dealer{index}"]["SaleFile"]["PaymentCycle"]
        inventory_cycle = DealerConfig[f"Dealer{index}"]["InventoryFile"]["PaymentCycle"]
        
        # 取得經銷商檔案副檔名
        sale_extension = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
        inventory_extension = DealerConfig[f"Dealer{index}"]["InventoryFile"]["Extension"]
        
        # 目錄無檔案的經銷商ID，加入List
        if not file_names:
            not_submission.append(dealer_id)
        
        for file_name in file_names:
            file_path = os.path.join(dealer_path, file_name)
            file_update_time = os.path.getmtime(file_path)
            file_update_time = datetime.fromtimestamp(file_update_time).date()
            file_type, data_max_row = decide_file_type(dealer_id, file_name)

            # 依據檔案類型切換參數
            if file_type is not None:
                msg = f"{file_type} 檔案準時繳交，繳交時間 {file_update_time}"
                WRecLog("1", "RecordDealerfiles", dealer_id, file_name, msg)
                if file_type == "Sale":
                    file_cycle = sale_cycle
                    file_extension = sale_extension
                elif file_type == "Inventory":
                    file_cycle = inventory_cycle
                    file_extension = inventory_extension
                
                # 取得檔名中的時間參數
                file_time = check_file_name_format(dealer_id, file_name, file_extension)
                
                # 寫入sub_record.json
                input_data = {f"Dealer{index}":{f"{file_type}File":True}}
                msg = SubRecordJson("WriteFileStatus", input_data)
                WRecLog("1", "RecordDealerFiles", dealer_id, file_name, msg)

                if file_cycle == "D":
                    start_time = file_time 
                    end_time = file_time + timedelta(days = 1)
                else:
                    start_time = SystemTime.date().replace(day = 1)
                    end_time = SystemTime.date().replace(day = MonthlyFileRange)
                time_due = f"{start_time} ~ {end_time}"
                if (start_time <= file_update_time) and (file_update_time <= end_time):
                    status = "有繳交"
                elif end_time < file_update_time:
                    status = "補繳交"
                else:
                    status = "時間錯誤"
            
            # 寫入繳交紀錄
            write_data = {"UploadData":{
                "經銷商ID":dealer_id,
                "檔案類型":file_type,
                "繳交狀態":status,
                "檔案名稱":file_name,
                "應繳時間":time_due,
                "繳交時間":file_update_time,
                "檔案內容總筆數":data_max_row
            }}
            result, data_id = WriteSubRawData(write_data)
            if result:
                data_dic[data_id] = file_name
    
    # 寫入未繳交紀錄
    write_not_sub_record()
    have_submission = list(set(DealerList) - set(not_submission))
    for dealer_id in not_submission:
        msg = "檔案未繳交"
        WRecLog("2", "RecordDealerFiles", dealer_id, None, msg)
    
    return have_submission, data_dic

# 清空sub_record.json
def ClearSubRecordJson():
    for i in range(len(DealerList)):
        index = i + 1
        sale_file_cycle = DealerConfig[f"Dealer{index}"]["SaleFile"]["PaymentCycle"]
        inventory_file_cycle = DealerConfig[f"Dealer{index}"]["InventoryFile"]["PaymentCycle"]
        for j in range(2):
            file_cycle = sale_file_cycle if j == 0 else inventory_file_cycle
            file_type = "Sale" if j == 0 else "Inventory"
            if file_cycle == "D":
                input_data = {f"Dealer{index}":{f"{file_type}File":None}}
                msg = SubRecordJson("WriteFileSatus", input_data)
                WSysLog("1", "ClearSubRecordJson", msg)
            else:
                next_day = SystemTime + timedelta(days = 1)
                if next_day.month != SystemTime.month:
                    input_data = {f"Dealer{index}":{f"{file_type}File":None}}
                    msg = SubRecordJson("WriteFileSatus", input_data)
                    WSysLog("1", "ClearSubRecordJson", msg)

# 搬移檢查出錯誤的檔案
def move_error_file(dealer_id, file_names):
    folder_name = SystemTime.strftime("%Y%m")
    source_path = os.path.join(DealerPath, dealer_id)
    target_path = os.path.join(source_path, CompletedDir, folder_name)
    if not os.path.exists(target_path):
        os.makedirs(target_path)
        msg = f"已在 {CompletedDir} 目錄下建立資料夾，資料夾名稱 {folder_name}"
        WSysLog("1", "MoveErrorFile", msg)

    for file_name in file_names:
        file_source = os.path.join(source_path, file_name)
        file_target = os.path.join(target_path, file_name)
        shutil.move(file_source, file_target)
        if os.path.exists(file_target):
            msg = f"檔案搬移至 {target_path} 成功"
            WSysLog("1", "MoveErrorFile", msg)
        else:
            msg = f"檔案搬移至 {target_path} 失敗"
            WSysLog("2", "MoveErrorFile", msg)

# 檢查檔案表頭是否符合
def CheckFileHeader(dealer_id, file_name, file_type):
    flag = False
    format_header = SF_Default_Header if file_type == "Sale" else IF_Default_Header
    must_have_header = SF_MustHave if file_type == "Sale" else IF_MustHave
    two_choose_one = SF_2Choose1 if file_type == "Sale" else IF_2Choose1
    file_path = os.path.join(DealerPath, dealer_id, file_name)
    file_data, _ = read_data(file_path)
    file_header = file_data.columns.tolist()
    if set(file_header) == set(format_header):
        msg = "全部表頭都存在"
        WCheLog("1", "CheckFileHeader", dealer_id, file_name, msg)
        flag = True
    else:
        if (must_have_header.issubset(file_header)) and\
            (two_choose_one.issubset(file_header)):
            msg = "必要表頭都存在"
            WCheLog("1", "CheckFileHeader", dealer_id, file_name, msg)
            flag = True
        else:
            less_header = list(set(must_have_header) - set(file_header))
            msg = f"必要表頭缺失，缺少表頭 {less_header}。"
            WCheLog("2", "CheckFileHeader", dealer_id, file_name, msg)
            file = os.path.splitext(file_name)[0]
            txt_path = os.path.join(DealerPath, dealer_id, f"{file}_header_error.txt")
            with open(txt_path, "w", encoding = "UTF-8") as error_txt:
                error_txt.write(msg)
            move_error_file(dealer_id, [file_name, f"{file}_header_error.txt"])

    if flag:
        return True
    else:
        return False

# 將col名稱轉變為excel內的名稱
def column_to_excel(col_index):
    """Convert a zero-based column index to Excel-style letter."""
    result = ""
    while col_index >= 0:
        result = chr(col_index % 26 + 65) + result
        col_index = col_index // 26 - 1
    return result

# 將 data 轉變為 excel 欄位
def change_to_excel_col_row(file_data, column, row):
    excel_row = row + 2
    excel_column_header = column_to_excel(file_data.columns.get_loc(column))
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

# 檢查檔案Creation Date日期是否符合
def check_date_time(dealer_id, file_name):
    flag = False
    cell_value, error_list = [], []
    file_path = os.path.join(DealerPath, dealer_id, file_name)
    file_update_time = os.path.getmtime(file_path)
    file_update_time = datetime.fromtimestamp(file_update_time).date()
    df, _ = read_data(file_path)
    df["Creation Date"] = pd.to_datetime(df['Creation Date'], format='%Y/%m/%d').dt.date
    df["is_valid"] = df['Creation Date'] <= file_update_time
    invalid_indices = df.index[~df["is_valid"]]
    for i in df["is_valid"]:
        if not i:
            flag = True

    if not flag:
        msg = "檔案時間符合。"
        WCheLog("1", "CheckDataTime", dealer_id, file_name, msg)
        return error_list, 0
    else:
        for i in invalid_indices:
            cell = change_to_excel_col_row(df, "Creation Date", i)
            cell_value.append(cell)
            msg = f"{cell} 與檔案更新時間不符合。"
            WCheLog("2", "CheckDataTime", dealer_id, file_name, msg)
        error_num = len(cell_value)
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])
        return error_list,  error_num

# 檢查檔案內容是否符合
def CheckFileContent(dealer_id, file_name, file_type):
    must_have_header = SF_MustHave if file_type == "Sale" else IF_MustHave
    two_choose_one = SF_2Choose1 if file_type == "Sale" else IF_2Choose1
    file_dir = os.path.join(DealerPath, dealer_id)
    file_path = os.path.join(DealerPath, dealer_id, file_name)
    file_data, _ = read_data(file_path)
    error_list, must_have_values = [], {}
    
    # 必要欄位值確認不為空
    for i in must_have_header:
        must_have_values[i] = file_data.index[file_data[i].isna()].tolist()
    
    cell_value = []
    for i in must_have_values:
        for j in must_have_values[i]:
            cell = change_to_excel_col_row(file_data, i ,j)
            cell_value.append(cell)
            msg = f"{cell} 內容為空。"
            WCheLog("2", "CheckFileContent", dealer_id, file_name ,msg)
    error_num = len(cell_value)
    if cell_value:
        cell_result = merge_ranges(cell_value)
        error_list.extend([f"{'、'.join(cell_result[i])} 內容為空" for i in cell_result])

    # 2選1選填欄位值確認不為空
    header1 = two_choose_one[0]
    header2 = two_choose_one[1]
    cell_value = []
    # Original Quantity欄位
    for i in file_data.index[file_data[header1].isna()].tolist():
        cell = change_to_excel_col_row(file_data, header1, i)
        msg = f"{cell} 內容為空"
        WCheLog("1", "CheckFileContent", dealer_id, file_name ,msg)

        # Quantity欄位
        if pd.isna(file_data.at[i, header2]):
            cell2 = change_to_excel_col_row(file_data, header2, i)
            msg = f"{cell} 及 {cell2} 內容皆為空"
            cell_value.append(cell)
            cell_value.append(cell2)
            WCheLog("2", "CheckFileContent", dealer_id, file_name ,msg)
        else:
            cell2 = change_to_excel_col_row(file_data, header2, i)
            msg = f"{cell2} 內容有值"
            WCheLog("1", "CheckFileContent", dealer_id, file_name ,msg)

    # 錯誤資訊統整
    error_num = error_num + len(cell_value)
    if cell_value:
        key = []
        cell_result = merge_ranges(cell_value)
        for i in cell_result:
            key.append(i)
        error_list.append\
        (f"{'、'.join(cell_result[key[0]])} 及 {'、'.join(cell_result[key[1]])} 內容為空")
    
    error_message, e_num = check_date_time(dealer_id, file_name)
    error_list += error_message
    error_num = error_num + e_num
    if not error_list:
        msg = "檔案內容正確"
        WCheLog("1", "CheckFileContent", dealer_id, file_name ,msg)
        return True, error_num
    else:
        file = os.path.splitext(file_name)[0]
        file_path = os.path.join(file_dir, f"{file}_content_error.txt")
        with open(file_path, "w", encoding = "UTF-8") as f:
            for i in range(len(error_list)):
                f.write(f"{i+1}. {error_list[i]}\n")
        move_error_file(dealer_id, [file_name, f"{file}_content_error.txt"])
        return False, error_num

# 檢查檔案主程式 ##
def CheckFile(have_submission, data_dic):
    for dealer_id in have_submission:
        folder_path = os.path.join(DealerPath, dealer_id)
        file_names = [file for file in os.listdir(folder_path) \
                      if os.path.isfile(os.path.join(folder_path, file))]
        for file_name in file_names:
            file_type, _ = decide_file_type(dealer_id, file_name)
            header_result = CheckFileHeader(dealer_id, file_name, file_type)
            if header_result:
                header_msg = "表頭正確"
                content_result, error_num = CheckFileContent(dealer_id, file_name, file_type)
                if content_result:
                    check_status = "OK"
                    content_msg = "內容正確"
                else:
                    check_status = "NO"
                    content_msg = "內容錯誤"
            else:
                check_status = "ON"
                header_msg = "表頭錯誤"
                content_msg = "無檢查"
                error_num = 0
            
            # 尋找檔案id值
            for data_id, file_name_in_tabel in data_dic.items():
                if file_name == file_name_in_tabel:
                    data_id = data_id
                    break
            
            # 寫入繳交紀錄
            write_data = {"CheckData":{
                "ID":data_id,
                "檢查狀態":check_status,
                "表頭檢查結果":header_msg,
                "內容檢查結果":content_msg,
                "內容錯誤筆數":error_num
            }}
            WriteSubRawData(write_data)

if __name__ == "__main__":
    sub_list, file_id, = RecordDealerFiles()
    print(f"file_id:{file_id}")
    CheckFile(sub_list)