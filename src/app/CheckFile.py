# -*- coding: utf-8 -*-

'''
檔案說明：確認檔案繳交時間，
檢查檔案副檔名、表頭格式及內容
Writer:Qian
'''

import os, re
import shutil
import pandas as pd
from Mail import SendMail
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

# 共用函數
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

# RecordDealerFiles
# 寫入未繳交名單
def write_not_sub_record(notify):
    mail_data = []
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
                "缺繳(待補繳)檔案名稱":f"{dealer_id}_S_{SystemTime.date().strftime('%Y%m%d')}.{file_extencion}",
                "檔案狀態":"未繳交",
                "應繳時間":f"{SystemTime.date()} ~ {SystemTime.date() + timedelta(days = 1)}",
                "檔案檢查結果":"未檢查"
            }
            if notify:
                mail_data.append("銷售檔案")
            WriteNotSubmission(write_data)
        
        # 庫存日繳
        if (inventory_cycle == "D") and (not record_data[f"Dealer{index}"]["InventoryFile"]):
            file_extencion = DealerConfig[f"Dealer{index}"]["InventoryFile"]["Extension"]
            write_data = {
                "經銷商ID":dealer_id,
                "檔案類型":"Inventory",
                "缺繳(待補繳)檔案名稱":f"{dealer_id}_I_{SystemTime.date().strftime('%Y%m%d')}.{file_extencion}",
                "檔案狀態":"未繳交",
                "應繳時間":f"{SystemTime.date()} ~ {SystemTime.date() + timedelta(days = 1)}",
                "檔案檢查結果":"未檢查"
            }
            if notify:
                mail_data.append("庫存檔案")
            WriteNotSubmission(write_data)
        
        # 銷售月繳
        if (sale_cycle == "M") and (SystemTime == (SystemTime.replace(day = MonthlyFileRange))):
            if not record_data[f"Dealer{index}"]["SaleFile"]:
                file_extencion = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
                write_data = {
                    "經銷商ID":dealer_id,
                    "檔案類型":"Sale",
                    "缺繳(待補繳)檔案名稱":f"{dealer_id}_S_{(SystemTime.date().replace(day = 1) - timedelta(days = 1)).strftime('%Y%m%d')}.{file_extencion}",
                    "檔案狀態":"未繳交",
                    "應繳時間":f"{SystemTime.date().replace(day = 1)} ~ {SystemTime.date().replace(day = MonthlyFileRange)}",
                    "檔案檢查結果":"未檢查"
                }
                if notify:
                    mail_data.append("銷售檔案")
                WriteNotSubmission(write_data)
        
        # 庫存月繳
        if (inventory_cycle == "M") and (SystemTime == (SystemTime.replace(day = MonthlyFileRange))):
            if not record_data[f"Dealer{index}"]["InventoryFile"]:
                file_extencion = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
                write_data = {
                    "經銷商ID":dealer_id,
                    "檔案類型":"Inventory",
                    "缺繳(待補繳)檔案名稱":f"{dealer_id}_I_{(SystemTime.date().replace(day = 1) - timedelta(days = 1)).strftime('%Y%m%d')}.{file_extencion}",
                    "檔案狀態":"未繳交",
                    "應繳時間":f"{SystemTime.date().replace(day = 1)} ~ {SystemTime.date().replace(day = MonthlyFileRange)}",
                    "檔案檢查結果":"未檢查"
                }
                if notify:
                    mail_data.append("庫存檔案")
                WriteNotSubmission(write_data)
        if notify:
            mail_data = "、".join(mail_data)
            send_info = {"Mode" : "FileNotSub", "DealerID" : dealer_id, "MailData" : mail_data, "FilesPath" : None}
            SendMail(send_info)

# 檢查檔案命名格式
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
            file_time = file_time.date()
        except ValueError as  e:
            flag = False
            msg = f"檔名內容錯誤，時間內容錯誤 {e}。"
            WRecLog("2", "RecordDealerfiles", dealer_id, file_name, msg)
    elif len(file_name_part2) == 12:
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
            os.remove(file_path)
            msg = f"已移除檔案 {file_name}。"
            WSysLog("1", "RecordDealerfiles", msg)
        except FileNotFoundError:
            msg = f"系統找不到該檔案 {file_name} ，無法移除。"
            WSysLog("2", "RecordDealerfiles", msg)
        except Exception as e:
            msg = f"移除檔案 {file_name} 時發生未知錯誤： {e}。"
            WSysLog("2", "RecordDealerfiles", msg)
    return file_time

# 整理待繳紀錄表中日繳檔案
def daily_file_resub(dealer_id, file_type, file_name):
    # 抓取紀錄中的經銷商副檔名
    for i in range(len(DealerList)):
        if dealer_id == DealerList[i]:
            index = i + 1
            break
    sale_extension = DealerConfig[f"Dealer{index}"]["SaleFile"]["Extension"]
    inventory_extension = DealerConfig[f"Dealer{index}"]["InventoryFile"]["Extension"]

    # 讀取待補繳紀錄表中篩選的內容
    header = GlobalConfig["NotSubmission"]["Header"]
    data = WriteNotSubmission("ReadDaily")
    # 針對日繳未繳交處理
    df = data[(data[header[1]] == dealer_id) & (data[header[3]] == file_type) & (data[header[6]] == "未繳交")]
    not_sub_file = df[header[5]].values
    not_sub_date_list = []
    for item in not_sub_file:
        name, _ = item.rsplit(".", 1)
        parts = re.split(r"_", name)
        not_sub_date = datetime.strptime(parts[2], "%Y%m%d")
        not_sub_date_list.append(not_sub_date)
    not_sub_date_list.sort()

    # 抓取目標時間
    target_file_dict = {}
    processed_dates = set()
    for date in not_sub_date_list:
        if date in processed_dates:
            continue
        source_dates = [date]
        next_date = date + timedelta(days=1)
        
        while next_date in not_sub_date_list:
            source_dates.append(next_date)
            processed_dates.add(next_date)
            next_date += timedelta(days=1)
        
        target_date = source_dates[-1] + timedelta(days=1)
        processed_dates.update(source_dates)
        target_file_dict[target_date] = source_dates

    # 生成檔案名稱字典，目標:來源
    final_dict = {}
    data_type = "S" if file_type == "Sale" else "I"
    extension = sale_extension if file_type == "Sale" else inventory_extension
    for target_date, source_dates in target_file_dict.items():
        target_file_name = f"{dealer_id}_{data_type}_{target_date.strftime('%Y%m%d')}.{extension}"
        source_files = [f"{dealer_id}_{data_type}_{date.strftime('%Y%m%d')}.{extension}" for date in source_dates]
        final_dict[target_file_name] = source_files

    for target, source in final_dict.items():
        if target == file_name:
            return source
    return None

# 確認補繳檔案存在於待補繳清單
def monthly_file_resub(dealer_id, file_type, file_name):
    # 讀取待補繳紀錄表中篩選的內容
    header = GlobalConfig["NotSubmission"]["Header"]
    data = WriteNotSubmission("ReadMonthly")
    df = data[(data[header[1]] == dealer_id) & (data[header[3]] == file_type)]
    not_sub_file = df[header[5]].values
    if file_name in not_sub_file:
        return True, None
    else:
        return False, not_sub_file

# 紀錄檔案繳交主程式，回傳 有檔案名單，繳交dic，補繳交dic
def RecordDealerFiles(notify):
    not_submission, sub_dic, sub, resub = [], {}, {}, {}
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
            
            # 日繳檔案僅變更檔案狀態，無檢查紀錄
            resub_files = daily_file_resub(dealer_id, file_type, file_name)
            if resub_files is not None:
                for file in resub_files:
                    write_data = {
                        "經銷商ID":dealer_id,
                        "檔案類型":file_type,
                        "缺繳(待補繳)檔案名稱":file,
                        "檔案狀態":"已補繳",
                        "補繳時間":file_update_time,
                        "補繳檢查結果":"無檢查"
                    }
                    WriteNotSubmission(write_data)

            #月繳補繳，檔名不符合的處理
            if (file_cycle == "M") and (status == "補繳交"):
                result, not_sub_files = monthly_file_resub(dealer_id, file_type, file_name)
                if not result:
                    mail_data = {"FileName" : file_name, "SubFile" : "、".join(not_sub_files)}
                    send_info = {"Mode":"FileReSubError", "DealerID" : dealer_id, "MailData" : mail_data, "FilesPath" : None}
                    SendMail(send_info)
                    msg = f"檔案狀態為 {status}，但未存在於待補繳清單中。"
                    WRecLog("2", "MonthlyFileReSub", dealer_id, file_name, msg)
                    os.remove(file_path)
                    if not os.path.exists(file_path):
                        msg = "系統已刪除該檔案。"
                        WRecLog("1", "MonthlyFileReSub", dealer_id, file_name, msg)
                    else:
                        msg = "系統刪除檔案時遇到未知問題。"
                        WRecLog("2", "MonthlyFileReSub", dealer_id, file_name, msg)
                    break

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

            # 寫入繳交字典及補繳交字典
            result, sub_data_id = WriteSubRawData(write_data)
            if result:
                sub_dic[sub_data_id] = file_name
                if status == "補繳交":
                    resub[file_name] = file_update_time
                else:
                    sub[file_name] = time_due

    # 寫入未繳交紀錄
    write_not_sub_record(notify)

    have_submission = list(set(DealerList) - set(not_submission))
    for dealer_id in not_submission:
        msg = "檔案未繳交"
        WRecLog("2", "RecordDealerFiles", dealer_id, None, msg)
    
    return have_submission, sub_dic, sub, resub

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
                #當月最後一天才刷新月繳的參數
                next_day = SystemTime + timedelta(days = 1)
                if next_day.month != SystemTime.month:
                    input_data = {f"Dealer{index}":{f"{file_type}File":None}}
                    msg = SubRecordJson("WriteFileSatus", input_data)
                    WSysLog("1", "ClearSubRecordJson", msg)

# Check共用
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

    return target_path

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
            error_file_path = move_error_file(dealer_id, [file_name, f"{file}_header_error.txt"])
            mail_data = {"FileName": file_name}
            mail_data_path = os.path.join(error_file_path, f"{file}_header_error.txt")
            send_info = {"Mode":"FileContentError", "DealerID": dealer_id, "MailData": mail_data, "FilesPath": mail_data_path}
            SendMail(send_info)
            
    if flag:
        return True
    else:
        return False

# CheckFileContent
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
        error_file_path = move_error_file(dealer_id, [file_name, f"{file}_content_error.txt"])
        mail_data = {"FileName": file_name}
        mail_data_path = os.path.join(error_file_path, f"{file}_header_error.txt")
        send_info = {"Mode":"FileContentError", "DealerID": dealer_id, "MailData": mail_data, "FilesPath": mail_data_path}
        SendMail(send_info)
        return False, error_num

# 檢查檔案主程式
def CheckFile(have_submission, sub_dic, sub, resub):
    change_dic = {}
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
                check_status = "NO"
                header_msg = "表頭錯誤"
                content_msg = "無檢查"
                error_num = 0
            
            # 一般繳交
            # 尋找檔案id值
            for data_id, file_name_in_tabel in sub_dic.items():
                if file_name == file_name_in_tabel:
                    data_index = data_id
                    break
            
            # 寫入繳交紀錄
            write_data = {"CheckData":{
                "ID":data_index,
                "檢查狀態":check_status,
                "表頭檢查結果":header_msg,
                "內容檢查結果":content_msg,
                "內容錯誤筆數":error_num
            }}
            WriteSubRawData(write_data)

            if check_status == "OK":
                change_dic[data_index] = file_name

            # 寫入待補繳紀錄表，檔案有交但有錯
            if check_status == "NO":
                for sub_file_name, sub_file_time_due in sub.items():
                    if file_name == sub_file_name:
                        write_data = {
                            "經銷商ID":dealer_id,
                            "檔案類型":file_type,
                            "缺繳(待補繳)檔案名稱":file_name,
                            "檔案狀態":"有繳交",
                            "應繳時間":sub_file_time_due,
                            "檔案檢查結果":f"{header_msg};{content_msg}"
                        }
                        WriteNotSubmission(write_data)

            # 檔案補繳
            # 將補繳檔案檢查結果寫入待補繳紀錄表中
            for resub_file_name, resub_file_upload_time in resub.items():
                if file_name == resub_file_name:
                    write_data = {
                        "經銷商ID":dealer_id,
                        "檔案類型":file_type,
                        "缺繳(待補繳)檔案名稱":file_name,
                        "檔案狀態":"已補繳",
                        "補繳時間":resub_file_upload_time,
                        "補繳檢查結果":f"{header_msg};{content_msg}"
                    }
                    WriteNotSubmission(write_data)
    if not change_dic:
        change_dic = None
    return change_dic
      
if __name__ == "__main__":
    HaveSubmission, SubDic, Sub, ReSub = RecordDealerFiles(True)
    ClearSubRecordJson()
    ChangeDic = CheckFile(HaveSubmission, SubDic, Sub, ReSub)
    print(ChangeDic)