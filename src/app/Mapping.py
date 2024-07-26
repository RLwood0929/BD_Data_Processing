# -*- coding: utf-8 -*-

'''
檔案說明：檔案進行格式轉換
Writer:Qian
'''
# 分類檔案類型，區分銷售及庫存檔案

# 庫存檔案由各經銷商資料轉換後，統一合併成一份檔案

# 轉換部分套用轉換規則，默認規則為搬移

import os
import pandas as pd
from datetime import datetime
from CheckFile import read_data
from SystemConfig import Config, MappingRule, DealerConf

GlobalConfig = Config()
DealerConfig = DealerConf()
MappingConfig = MappingRule()

# 全域目錄參數
DealerList = DealerConfig["DealerList"]
KADealerList = DealerConfig["KADealerList"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
ChangeFileDir = GlobalConfig["Default"]["ChangeFileDir"]
BDFolderDir = GlobalConfig["Default"]["BDFolderName"]
MasterFileDir = GlobalConfig["Default"]["MasterFileDir"]

DealerPath = os.path.join(RootDir, FolderName, DealerDir)
ChangedPath = os.path.join(DealerPath, ChangeFileDir)
MasterFolderPath = os.path.join(RootDir, FolderName, BDFolderDir, MasterFileDir)
MasterFile = [file for file in os.listdir(MasterFolderPath) \
            if os.path.isfile(os.path.join(MasterFolderPath, file))]

SaleFileChangeRule = MappingConfig["MappingRule"]["Sale"]
InventoryFileChangeRule = MappingConfig["MappingRule"]["Inventory"]

# 銷售輸出參數
SaleOutputFileName = GlobalConfig["OutputFile"]["Sale"]["FileName"]
SaleOutputFileHeader = GlobalConfig["OutputFile"]["Sale"]["Header"]
SaleOutputFileExtension = GlobalConfig["OutputFile"]["Sale"]["Extension"]
SaleErrorReportFileName = GlobalConfig["SaleErrorReport"]["FileName"]
SaleErrorReportHeader = GlobalConfig["SaleErrorReport"]["Header"]

# 庫存輸出參數
InventoryOutputFileName = GlobalConfig["OutputFile"]["Inventory"]["FileName"]
InventoryOutputFileHeader = GlobalConfig["OutputFile"]["Inventory"]["Header"]
InventoryOutputFileExtension = GlobalConfig["OutputFile"]["Inventory"]["Extension"]
InventoryOutputFileCountryCode = GlobalConfig["OutputFile"]["Inventory"]["CountryCode"]
InventoryOutputFileName = InventoryOutputFileName.replace("{CountryCode}", InventoryOutputFileCountryCode)
InventoryErrorReportFileName = GlobalConfig["InventoryErrorReport"]["FileName"]
InventoryErrorReportHeader = GlobalConfig["InventoryErrorReport"]["Header"]

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
    return True, master_data, ka_data

# 比對、篩選 product id 不存在於 masterfile 中的資料 #
def check_product_id(dealer_id, input_data):
    result, master_data, _ = read_master_file()
    error_data = pd.DataFrame(columns = input_data.columns)
    if result:
        master_col = master_data.columns.values
        search_data = master_data[master_data[master_col[0]] == dealer_id]
        error_row = []
        for index, row in input_data.iterrows():
            product_id = str(row["Product ID"])
            if product_id not in search_data[master_col[1]].values:
                print()
        print(error_data)
        
# Quantity特殊處理
def move_or_search_uom(input_data, source_col, target_col, dealer_id):
    output, error_row = [], []
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
                    error_row.append(row)
                else:
                    uom = uom_list[-1]
                    output.append(int(uom) * int(input_data[target_col][row]))
        return output, error_row

# 使用 product id 在 MasterFile 中找到對應的 DP 價
def search_dp(input_data, dealer_id):
    result, master_data, ka_data = read_master_file()
    output, error_row = [], []
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
                        error_row.append(row)
                        msg = f"MasterFile檔案中， {product_id} 該 {price_type} 值為空。"
            if not in_range_flag:
                msg = f"{product_id} 在MasterFile檔案中，未搜尋到起迄區間符合 {data_date} 之資料。"
                error_row.append(row)
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

# 依據轉換規則轉換銷售檔案 #
def ChangeSaleFile(dealer_id, file_path):
    file_header = SaleOutputFileHeader
    change_rules = SaleFileChangeRule
    input_data, data_max_row = read_data(file_path)
    output_data = pd.DataFrame(columns = file_header)
    check_product_id(dealer_id, input_data)
    for rule_index in range(1, len(change_rules) + 1):
        source_col = change_rules[f"Column{rule_index}"]["SourceName"]
        target_col = change_rules[f"Column{rule_index}"]["ColumnName"]
        rule = change_rules[f"Column{rule_index}"]["ChangeRule"]
        value = change_rules[f"Column{rule_index}"]["Value"]
        # 欄位為固定值
        if rule == "FixedValue":
            if target_col == "Area":
                value = "TWM" #
            output_data[target_col] = fixed_value(value, data_max_row)
        # 搬移資料
        elif rule == "Move":
            output_data[target_col] = move_rule(input_data, source_col)
        # 變更時間格式
        elif rule == "ChangeTimeFormat":
            output_data[target_col] = change_time_format(input_data, source_col, value)
        # Quantity特殊處理
        elif rule == "MoveOrSearchUom":
            output, error_row = move_or_search_uom(input_data, source_col, target_col, dealer_id)
            if error_row:
                output_data = output_data.drop(error_row).reset_index(drop=True)
            output_data[target_col] = output
        # 搜索MasterFile的dp價
        elif rule == "SearchDP":
            output, error_row = search_dp(input_data, dealer_id)
            if error_row:
                output_data = output_data.drop(error_row).reset_index(drop=True)
            output_data[target_col] = output
        # 多欄位內容合併
        elif rule == "MergeColumns":
            output_data[target_col] = merge_columns(input_data, source_col, value)
        elif not rule:
            continue
        else:
            msg = f"{rule} 此轉換規則不再範圍中。"
    
    # 輸出傳換後的sale檔案
    data_start_date = datetime.strftime(input_data["Transaction Date"][0], "%Y%m%d")
    data_end_date = datetime.strftime(input_data["Transaction Date"][data_max_row - 1], "%Y%m%d")
    file_name = SaleOutputFileName.replace("{DealerID}", str(dealer_id))\
        .replace("{DataStartDate}", data_start_date).replace("{DataEndDate}", data_end_date)
    output_data.to_csv(os.path.join(ChangedPath, dealer_id, f"{file_name}.{SaleOutputFileExtension}"), index=False)

# 新欄位內容填寫 Transaction Date 最後一天
def last_transaction_date(input_data, row):
    return [str(datetime.strftime(input_data["Transaction Date"][row - 1], "%Y%m%d"))] * row

# 依據規則轉換庫存檔案 #
def ChangeInventoryFile(dealer_id, file_path):
    file_header = InventoryOutputFileHeader
    change_rules = InventoryFileChangeRule
    input_data, data_max_row = read_data(file_path)
    output_data = pd.DataFrame(columns = file_header)
    check_product_id(dealer_id, input_data)
    for rule_index in range(1, len(change_rules) + 1):
        source_col = change_rules[f"Column{rule_index}"]["SourceName"]
        target_col = change_rules[f"Column{rule_index}"]["ColumnName"]
        rule = change_rules[f"Column{rule_index}"]["ChangeRule"]
        value = change_rules[f"Column{rule_index}"]["Value"]
        # 欄位為固定值
        if rule == "FixedValue":
            output_data[target_col] = fixed_value(value, data_max_row)
        # 搬移資料
        elif rule == "Move":
            output_data[target_col] = move_rule(input_data, source_col)
        # Quantity特殊處理
        elif rule == "MoveOrSearchUom":
            output, error_row = move_or_search_uom(input_data, source_col, target_col, dealer_id)
            if error_row:
                output_data = output_data.drop(error_row).reset_index(drop=True)
            output_data[target_col] = output
        # Date Period欄位為 Transaction Date 最後一天
        elif rule == "LastTransactionDate":
            output_data[target_col] = last_transaction_date(input_data, data_max_row)
    last_month_day = datetime.strftime(input_data["Transaction Date"][data_max_row - 1], "%Y%m%d")
    file_name = InventoryOutputFileName.replace("{CountryCode}", InventoryOutputFileCountryCode)\
        .replace("{LastTransactionDate}", last_month_day)
    output_data.to_csv(os.path.join(ChangedPath, dealer_id, f"{file_name}.{InventoryOutputFileExtension}"), sep = "\t", index=False)

# log需加入
# 待寫轉換主程式

if __name__ == "__main__":
    dealerID = "111"
    FilePath = os.path.join(DealerPath, "111/111_S_20240725.xlsx")
    # ChangeSaleFile(dealerID, FilePath)
    ChangeInventoryFile(dealerID, FilePath)