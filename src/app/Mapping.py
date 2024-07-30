# -*- coding: utf-8 -*-

'''
檔案說明：檔案進行格式轉換
Writer:Qian
'''
# 需加入mail
import os, re
import shutil
import pandas as pd
from datetime import datetime
from Log import WSysLog, WChaLog
from CheckFile import decide_file_type
from RecordTable import WriteSubRawData
from SystemConfig import Config, MappingRule, DealerConf

SystemTime = datetime.now()
GlobalConfig = Config()
DealerConfig = DealerConf()
MappingConfig = MappingRule()

DealerList = DealerConfig["DealerList"]
KADealerList = DealerConfig["KADealerList"]

# 全域目錄參數
RootDir = GlobalConfig["DirTree"]["Path"]
DealerDir = GlobalConfig["DirTree"]["Dealer"]["FolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
ChangeFileDir = GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["ChangeFileFolder"]["FolderName"]
BDFolderDir = GlobalConfig["DirTree"]["BD"]["FolderName"]
MasterFileDir = GlobalConfig["DirTree"]["BD"]["NextFolder"]["MasterFileFolder"]

DealerPath = os.path.join(RootDir, FolderName, DealerDir)
ChangedPath = os.path.join(DealerPath, ChangeFileDir)
MasterFolderPath = os.path.join(RootDir, FolderName, BDFolderDir, MasterFileDir)
MasterFile = [file for file in os.listdir(MasterFolderPath) \
            if os.path.isfile(os.path.join(MasterFolderPath, file))]
ErrorReportPath = os.path.join()

SaleFileChangeRule = MappingConfig["MappingRule"]["Sale"]
InventoryFileChangeRule = MappingConfig["MappingRule"]["Inventory"]

# 銷售輸出參數
SaleOutputFileName = GlobalConfig["OutputFile"]["Sale"]["FileName"]
SaleOutputFileHeader = GlobalConfig["OutputFile"]["Sale"]["Header"]
SaleOutputFileExtension = GlobalConfig["OutputFile"]["Sale"]["Extension"]
SaleErrorReportFileName = GlobalConfig["ErrorReport"]["Sale"]["FileName"]
SaleErrorReportHeader = GlobalConfig["ErrorReport"]["Sale"]["Header"]

# 庫存輸出參數
InventoryOutputFileName = GlobalConfig["OutputFile"]["Inventory"]["FileName"]
InventoryOutputFileHeader = GlobalConfig["OutputFile"]["Inventory"]["Header"]
InventoryOutputFileExtension = GlobalConfig["OutputFile"]["Inventory"]["Extension"]
InventoryOutputFileCountryCode = GlobalConfig["OutputFile"]["Inventory"]["CountryCode"]
InventoryOutputFileName = InventoryOutputFileName.replace("{CountryCode}", InventoryOutputFileCountryCode)
InventoryErrorReportFileName = GlobalConfig["ErrorReport"]["Inventory"]["FileName"]
InventoryErrorReportHeader = GlobalConfig["ErrorReport"]["Inventory"]["Header"]

# 讀取檔案
def read_data(file_path):
    file = os.path.basename(file_path)
    _, file_extension = os.path.splitext(file)
    file_extension = file_extension.lower()
    if file_extension == ".csv":
        df = pd.read_csv(file_path, dtype = str)
        return df
    elif file_extension in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, dtype = str)
        return df

# 搬移欄位值至新欄位
def move_rule(input_data, input_col):
    return input_data[input_col]

# 新欄位為固定值
def fixed_value(value, row):
    return [value]*row

# 轉換欄位內容的時間格式
def change_time_format(input_data, input_col, date_format):
    return input_data[input_col].dt.strftime(date_format)

# 讀取MasterFile
def read_master_file():
    if len(MasterFile) != 1:
        msg = f"{MasterFolderPath} 目標路徑下存在多份MasterFile，系統無法辨別使用何MasterFile。"
        return False, None, None
    master_file_path = os.path.join(MasterFolderPath, MasterFile[0])
    master_data = pd.read_excel(master_file_path,sheet_name = "MasterFile", dtype = str)
    ka_data = pd.read_excel(master_file_path,sheet_name = "KAList", dtype = str)
    msg = "成功讀取 MasterFile 資料。"
    WSysLog("1", "ReadMasterFile", msg)
    return True, master_data, ka_data

# 比對、篩選 product id 不存在於 masterfile 中的資料
def check_product_id(dealer_id, input_data):
    for i in range(len(DealerList)):
        if DealerList[i] == dealer_id:
            index = i + 1
            break
    dealer_name = DealerConfig[f"Dealer{index}"]["DealerName"]
    result, master_data, _ = read_master_file()
    error_data = pd.DataFrame(columns = input_data.columns)
    error_index = []
    if result:
        master_col = master_data.columns.values
        search_data = master_data[master_data[master_col[0]] == dealer_id]
        masterfile_product = set(search_data[master_col[1]].values)
        data_product = {str(pid) for pid in (set(input_data["Product ID"].to_list()))}
        not_in_masterfile = list(data_product - masterfile_product)
        for product_id in not_in_masterfile:
            error_row = input_data[input_data["Product ID"] == product_id]
            error_index.extend(error_row.index.values.tolist())
            error_index.sort()
            for row in range(len(error_row)):
                row_data = error_row.iloc[row].to_dict()
                error_data.loc[len(error_data)] = row_data
        error_data.insert(0,"Dealer ID", dealer_id)
        error_data.insert(1,"Dealer Name", dealer_name)
        error_msg = "Masterfile 檔案中無此 Product ID。"
        error_data.insert(2,"Exchange Error Issue", error_msg)
        return error_data, error_index
    
# Quantity特殊處理
def move_or_search_uom(input_data, source_col, target_col, dealer_id):
    output, error_row = [], {}
    result, master_data, _ = read_master_file()
    if result:
        master_col = master_data.columns.values
        for row in range(len(input_data)):
            source_value = input_data[source_col][row]
            if pd.notna(source_value):
                output.append(str(source_value))
            else:
                # 從經銷商檔案中取得產品id及日期
                product_id = str(input_data["Product ID"][row])
                data_date = input_data["Transaction Date"][row]

                # 讀取並篩選 masterfile 中對應的資料
                search_data = master_data[(master_data[master_col[0]] == dealer_id) & \
                                        (master_data[master_col[1]] == product_id)]
                search_data = search_data.reset_index(drop=True)
                in_range_flag = False
                uom_list = []
                for i in range(len(search_data)):
                    start_date = datetime.strptime(search_data[master_col[7]][i], "%Y%m%d")
                    end_date = datetime.strptime(search_data[master_col[8]][i], "%Y%m%d")
                    if start_date <= data_date <= end_date:
                        in_range_flag = True
                        uom = search_data[master_col[2]][i]
                        uom_list.append(uom)
                if not in_range_flag:
                    msg = f"{product_id} 在MasterFile檔案中，未搜尋到起迄區間符合 {data_date} 之資料。"
                    error_row[row] = msg
                else:
                    uom = uom_list[-1]
                    output.append(int(uom) * int(input_data[target_col][row]))
        return output, error_row

# 使用 product id 在 MasterFile 中找到對應的 DP 價
def search_dp(input_data, dealer_id):
    result, master_data, ka_data = read_master_file()
    output, error_row = [], {}
    if result:
        master_col = master_data.columns.values
        ka_col = ka_data.columns.values
        price_type = master_col[4]
        for row in range(len(input_data)):
            product_id = str(input_data["Product ID"][row])
            data_date = input_data["Transaction Date"][row]
            search_data = master_data[(master_data[master_col[0]] == dealer_id) & \
                                        (master_data[master_col[1]] == product_id)]
            search_data = search_data.reset_index(drop=True)

            # 針對KA經銷商進行DP價判斷
            if dealer_id in KADealerList:
                buyer = input_data["Buyer ID"][row]
                search_ka_data = ka_data[(ka_data[ka_col[0]] == dealer_id) &\
                                         (ka_data[ka_col[1]] == buyer)]
                ka_range = False
                price_type = []
                for i in range(len(search_ka_data)):
                    start_date = datetime.strptime(ka_data[ka_col[2]][i], "%Y%m%d")
                    end_date = datetime.strptime(ka_data[ka_col[3]][i], "%Y%m%d")
                    if start_date <= data_date <= end_date:
                        ka_range = True
                        ptype = ka_data[ka_col[4]][i]
                        price_type.append(ptype)
                if not ka_range:
                    price_type = master_col[4]
                else:
                    price_type = price_type[-1]

            in_range_flag = False
            dp_list = []
            for i in range(len(search_data)):
                start_date = datetime.strptime(search_data[master_col[7]][i], "%Y%m%d")
                end_date = datetime.strptime(search_data[master_col[8]][i], "%Y%m%d")
                if start_date <= data_date <= end_date:
                    in_range_flag = True
                    dp = search_data[price_type][i]
                    if pd.notna(dp):
                        dp_list.append(dp)
                    else:
                        msg = f"MasterFile檔案中， {product_id} 該 {price_type} 值為空。"
                        error_row[row] = msg
            if not in_range_flag:
                msg = f"{product_id} 在MasterFile檔案中，未搜尋到起迄區間符合 {data_date} 之資料。"
                error_row[row] = msg
            else:
                dp = dp_list[-1]
                output.append(dp)
        return output, error_row

# 多欄位值合併
def merge_columns(input_data, source_col, value):
    parts = source_col.split("+")
    output = input_data[parts].apply\
            (lambda row: value.join\
            (row.values.astype(str)), axis=1)
    return output

# 依據轉換規則轉換銷售檔案
def ChangeSaleFile(dealer_id, file_name):
    change_status = "OK"
    file_header = SaleOutputFileHeader
    change_rules = SaleFileChangeRule
    for i in range(len(DealerList)):
        if DealerList[i] == dealer_id:
            index = i + 1
            break

    dealer_name = DealerConfig[f"Dealer{index}"]["DealerName"]
    dealer_country = DealerConfig[f"Dealer{index}"]["Country"]
    file_path = os.path.join(DealerPath, dealer_id, file_name)
    input_data = read_data(file_path)

    input_data["Transaction Date"] = pd.to_datetime(input_data["Transaction Date"], format = "%Y/%m/%d")
    input_data["Creation Date"] = pd.to_datetime(input_data["Creation Date"], format = "%Y/%m/%d")
    output_data = pd.DataFrame(columns = file_header)
    error_data, error_index = check_product_id(dealer_id, input_data)
    msg = f"檔案中有 {len(error_data)} 筆資料在 master file 貨號中找不到。"
    WChaLog("2","ChangeSaleFile", dealer_id, file_name, msg)

    for i in error_index:
        input_data = input_data.drop(i)
    input_data = input_data.reset_index(drop = True)
    for rule_index in range(1, len(change_rules) + 1):
        source_col = change_rules[f"Column{rule_index}"]["SourceName"]
        target_col = change_rules[f"Column{rule_index}"]["ColumnName"]
        rule = change_rules[f"Column{rule_index}"]["ChangeRule"]
        value = change_rules[f"Column{rule_index}"]["Value"]
        # 欄位為固定值
        if rule == "FixedValue":
            if target_col == "Area":
                value = dealer_country
            output_data[target_col] = fixed_value(value, len(input_data))
        # 搬移資料
        elif rule == "Move":
            output_data[target_col] = move_rule(input_data, source_col)
        # 變更時間格式
        elif rule == "ChangeTimeFormat":
            output_data[target_col] = change_time_format(input_data, source_col, value)
        # Quantity特殊處理
        elif rule == "MoveOrSearchUom":
            output, error_row = move_or_search_uom(input_data, source_col, target_col, dealer_id)
            for row, msg in error_row.items():
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
        # 搜索MasterFile的dp價
        elif rule == "SearchDP":
            output, error_row = search_dp(input_data, dealer_id)
            for row, msg in error_row.items():
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
        # 多欄位內容合併
        elif rule == "MergeColumns":
            output_data[target_col] = merge_columns(input_data, source_col, value)
        elif not rule:
            continue
        else:
            change_status = "NO"
            msg = f"{rule} 此轉換規則不再範圍中。"
            WSysLog("3", "ChangeSaleFile", msg)

    # 輸出傳換後的sale檔案
    data_start_date = datetime.strftime(input_data["Transaction Date"][0], "%Y%m%d")
    data_end_date = datetime.strftime(input_data["Transaction Date"][len(input_data)- 1], "%Y%m%d")
    changed_file_name = SaleOutputFileName.replace("{DealerID}", str(dealer_id))\
        .replace("{TransactionDataStartDate}", data_start_date)\
        .replace("{TransactionDataEndDate}", data_end_date)
    try:
        output_data.to_csv(os.path.join(ChangedPath,\
            f"{changed_file_name}.{SaleOutputFileExtension}"), index=False)
        msg = f"檔案轉換完成，輸出檔名 {changed_file_name}.{SaleOutputFileExtension}。"
        WChaLog("1", "ChangeSaleFile", dealer_id, file_name, msg)
        if error_data:
            error_file_name = SaleErrorReportFileName.replace("{DealerID}", str(dealer_id))\
                .replace("{Date}", datetime.strftime(SystemTime, "%Y%m%d"))
            error_data.to_csv(os.path.join(ChangedPath, error_file_name), index = False)
            msg = f"Error檔案輸出完成，輸出檔名 {error_file_name}。"
            WChaLog("1", "ChangeSaleFile", dealer_id, file_name, msg)
        return change_status, f"{changed_file_name}.{SaleOutputFileExtension}",\
            len(error_data), len(output_data)
    except Exception as e:
        change_status = "NO"
        msg = f"轉換檔案失敗，遇到未知錯誤，{e}。"
        WChaLog("2", "ChangeSaleFile", dealer_id, file_name, msg)
        return change_status, None, 0, 0

# 新欄位內容填寫 Transaction Date 最後一天
def last_transaction_date(input_data, row):
    return [str(datetime.strftime(input_data["Transaction Date"][row - 1], "%m/%d/%Y"))] * row

# 依據規則轉換庫存檔案
def ChangeInventoryFile(dealer_id, file_name):
    change_status = "OK"
    file_header = InventoryOutputFileHeader
    change_rules = InventoryFileChangeRule
    for i in range(len(DealerList)):
        if DealerList[i] == dealer_id:
            index = i + 1
            break
    dealer_name = DealerConfig[f"Dealer{index}"]["DealerName"]
    file_path = os.path.join(DealerPath, dealer_id, file_name)
    input_data = read_data(file_path)
    input_data["Transaction Date"] = pd.to_datetime(input_data["Transaction Date"], format = "%Y/%m/%d")
    input_data["Creation Date"] = pd.to_datetime(input_data["Creation Date"], format = "%Y/%m/%d")
    output_data = pd.DataFrame(columns = file_header)
    error_data, error_index = check_product_id(dealer_id, input_data)
    for i in error_index:
        input_data = input_data.drop(i)
    input_data = input_data.reset_index(drop = True)
    for rule_index in range(1, len(change_rules) + 1):
        source_col = change_rules[f"Column{rule_index}"]["SourceName"]
        target_col = change_rules[f"Column{rule_index}"]["ColumnName"]
        rule = change_rules[f"Column{rule_index}"]["ChangeRule"]
        value = change_rules[f"Column{rule_index}"]["Value"]
        # 欄位為固定值
        if rule == "FixedValue":
            if source_col == "DealerID":
                value = dealer_id
            output_data[target_col] = fixed_value(value, len(input_data))
        # 搬移資料
        elif rule == "Move":
            output_data[target_col] = move_rule(input_data, source_col)
        # Quantity特殊處理
        elif rule == "MoveOrSearchUom":
            output, error_row = move_or_search_uom(input_data, source_col, target_col, dealer_id)
            for row, msg in error_row.items():
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
        # Date Period欄位為 Transaction Date 最後一天
        elif rule == "LastTransactionDate":
            output_data[target_col] = last_transaction_date(input_data, len(input_data))

    changed_file_name = f"{dealer_id}_I_{datetime.strftime(input_data['Transaction Date'][len(input_data) - 1], '%Y%m%d')}.csv"
    try:
        output_data.to_csv(os.path.join(ChangedPath, changed_file_name), index=False)
        msg = f"檔案轉換完成，輸出檔名 {changed_file_name}。"
        WChaLog("1", "ChangeInventoryFile", dealer_id, file_name, msg)
        if error_data:
            error_file_name = InventoryErrorReportFileName.replace("{DealerID}", str(dealer_id))\
                .replace("{Date}", datetime.strftime(SystemTime, "%Y%m%d"))
            error_data.to_csv(os.path.join(ChangedPath, error_file_name), index = False)
            msg = f"Error檔案輸出完成，輸出檔名 {error_file_name}。"
            WChaLog("1", "ChangeInventoryFile", dealer_id, file_name, msg)
        return change_status, changed_file_name,\
            len(error_data), len(output_data)
    except Exception as e:
        change_status = "NO"
        msg = f"轉換檔案失敗，遇到未知錯誤，{e}。"
        WChaLog("2", "ChangeInventoryFile", dealer_id, file_name, msg)
        return change_status, None, 0, 0

# 轉換主程式
def Changing(check_right_list):
    for dealer_id in DealerList:
        dealer_path = os.path.join(DealerPath, dealer_id)
        file_names = [file for file in os.listdir(dealer_path)\
                      if os.path.isfile(os.path.join(dealer_path, file))]
        for file_name in file_names:
            file_type, _ = decide_file_type(dealer_id, file_name)
            if file_type == "Sale":
                status, output_file_name, error_num, num = \
                    ChangeSaleFile(dealer_id, file_name)
            else:
                status, output_file_name, error_num, num = \
                    ChangeInventoryFile(dealer_id, file_name)
            for data_id, name in check_right_list:
                if file_name == name:
                    write_data = {
                        "ChangeData":{
                            "ID":data_id,
                            "轉換狀態":status,
                            "轉換後檔案名稱":output_file_name,
                            "轉換錯誤筆數":error_num,
                            "轉換後總筆數":num
                        }
                    }
                    WriteSubRawData(write_data)
                    break

# 合併 Inventory 檔案
def MargeInventoryFile():
    file_names = [file for file in os.listdir(ChangedPath) \
                      if os.path.isfile(os.path.join(ChangedPath, file))]
    
    # 取得要合併的檔案
    file_list, time_list = [], []
    for file_name in file_names:
        part = re.split(r"[._]" ,file_name)
        if part[1] == "I":
            file_list.append(file_name)
            time_list.append(part[2])

    # 抓取檔名中的檔案時間
    time_list = [datetime.strptime(time_str, "%Y%m%d") for time_str in time_list]
    file_time = time_list[-1].strftime("%Y%m%d")
    
    # 合併檔案
    dataframes = [pd.read_csv(os.path.join(ChangedPath, file)) for file in file_list]
    combined_df = pd.concat(dataframes, ignore_index=True)
    
    # 輸出檔案
    changed_file_name = InventoryOutputFileName.replace("{CountryCode}", InventoryOutputFileCountryCode)\
        .replace("{LastTransactionDate}", file_time)
    try:
        combined_df.to_csv(os.path.join(ChangedPath, \
            f"{changed_file_name}.{InventoryOutputFileExtension}"), sep = "\t", index=False)
        msg = f"{len(file_list)} 份檔案成功合併成 {changed_file_name}.{InventoryOutputFileExtension}。"
        WSysLog("1", "MargeInventoryFile", msg)
        for dealer_id in DealerList:
            target_folder = os.path.join(ChangedPath, dealer_id, datetime.strftime(SystemTime, "%Y%m"))
            for file in file_list:
                part = re.split(r"[._]" ,file)
                if (part[1] == dealer_id) and (part[2] == "I"):
                    file_source = os.path.join(ChangedPath, file)
                    file_target = os.path.join(target_folder, file)
                    shutil.move(file_source, file_target)
                    if os.path.exists(file_target):
                        msg = f"檔案搬移至 {target_folder} 成功"
                        WSysLog("1", "MoveInventoryFile", msg)
                    else:
                        msg = f"檔案搬移至 {target_folder} 失敗"
                        WSysLog("2", "MoveInventoryFile", msg)

    except Exception as e:
        msg = f"合併檔案發生未知錯誤，{e}。"
        WSysLog("3", "MargeInventoryFile", msg)

if __name__ == "__main__":
    aa = {22: '111_S_20240724.csv', 23: '111_S_20240725.xlsx', 24: '222_S_20240630.csv'}
    Changing(aa)
    # ChangeSaleFile(dealerID, FilePath)
    # input_data, data_max_row = read_data(FilePath)
    # dealerID = "111"
    # FilePath = "111_I_20240726.csv"
    # ChangeInventoryFile(dealerID, FilePath)
    # check_product_id(dealerID, input_data)
    # MargeInventoryFile()