# -*- coding: utf-8 -*-

'''
檔案說明：檔案進行格式轉換
Writer:Qian
'''

# 標準庫
import os, re, shutil
from datetime import datetime
from dateutil import parser

# 第三方庫
import pandas as pd

# 自定義函數
from Mail import SendMail
from Config import AppConfig
from Log import WSysLog, WChaLog
from CheckFile import get_file_type
from RecordTable import WriteSubRawData

Config = AppConfig()

# 統一時間欄位內容格式
def parse_and_format_date(date_str, output_format = "%Y/%m/%d"):
    try:
        parsed_date = parser.parse(date_str)
        return parsed_date.strftime(output_format)
    except (ValueError, TypeError):
        return date_str

# 讀取檔案
def read_data(file_path):
    file = os.path.basename(file_path)
    _, file_extension = os.path.splitext(file)
    file_extension = file_extension.lower()

    if file_extension in Config.AllowFileExtensions:
        df = pd.read_csv(file_path, dtype = str) \
                if file_extension == ".csv" \
                else pd.read_excel(file_path, dtype = str)
        df["Transaction Date"] = df["Transaction Date"].apply(lambda x: parse_and_format_date(str(x)))
        df["Transaction Date"] = pd.to_datetime(df["Transaction Date"], format = "%Y/%m/%d")
        df["Creation Date"] = df["Creation Date"].apply(lambda x: parse_and_format_date(str(x)))
        df["Creation Date"] = pd.to_datetime(df["Creation Date"], format = "%Y/%m/%d")
        # 過濾 Product ID 僅允許 a-z, A-Z, 0-9, - 符號
        df = df[df["Product ID"].str.contains("^[a-zA-Z0-9\-]+$", regex=True, na=False)]
        # 刷新index值
        df = df.reset_index()
        return df
    else:
        return False

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
    files = []
    if os.path.exists(Config.MasterFolderPath):
        files = [file for file in os.listdir(Config.MasterFolderPath) \
                if os.path.isfile(os.path.join(Config.MasterFolderPath, file))]
    if len(files) != 1:
        msg = f"{Config.MasterFolderPath} 目標路徑下存在多份MasterFile，系統無法辨別使用何MasterFile。"
        return False, None, None
    
    master_file_path = os.path.join(Config.MasterFolderPath, files[0])
    master_data = pd.read_excel(master_file_path,sheet_name = "MasterFile", dtype = str)
    ka_data = pd.read_excel(master_file_path,sheet_name = "KAList", dtype = str)
    msg = "成功讀取 MasterFile 資料。"
    WSysLog("1", "ReadMasterFile", msg)
    return True, master_data, ka_data

# 比對、篩選 product id 不存在於 masterfile 中的資料
def check_product_id(dealer_id, input_data):
    for i in range(len(Config.DealerList)):
        if Config.DealerList[i] == dealer_id:
            index = i + 1
            break

    dealer_name = Config.DealerConfig[f"Dealer{index}"]["DealerName"]
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
    print("Function:move_or_search_uom")
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
                    msg = f"此貨號未搜尋到起迄區間符合 {datetime.strftime(data_date, '%Y/%m/%d')} 之資料。"
                    error_row[row] = msg

                else:
                    uom = uom_list[-1]
                    output.append(int(uom) * int(input_data[target_col][row]))

        msg = f"ErrorRow結果： {error_row} 。"
        WSysLog("1", "MoveOrSearchUoM", msg)
        return output, error_row

# 使用 product id 在 MasterFile 中找到對應的 DP 價
def search_dp(input_data, dealer_id):
    print("Function: search_dp")
    result, master_data, ka_data = read_master_file()
    output, error_row, ptype_error = [], {}, {}

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
            if dealer_id in Config.KADealerList:
                buyer = input_data["Buyer ID"][row]
                search_ka_data = ka_data[(ka_data[ka_col[0]] == dealer_id) &\
                                         (ka_data[ka_col[1]] == buyer)]
                ka_range = False
                price_type = []

                for i in range(len(search_ka_data)):
                    start_date = datetime.strptime(ka_data[ka_col[2]][i], "%Y%m%d")
                    end_date = datetime.strptime(ka_data[ka_col[3]][i], "%Y%m%d")
                    data_date = data_date.to_pydatetime()

                    if start_date <= data_date <= end_date:
                        ka_range = True
                        ptype = ka_data[ka_col[4]][i]

                        if ptype == "KA":
                            ptype = master_col[5]

                        elif ptype == "IVY":
                            ptype = master_col[6]

                        price_type.append(ptype)

                    else:
                        msg = f"在ka表中， {dealer_id} 的 {buyer} 無符合起迄區間的資料。"
                        WSysLog("2", "SearchDP", msg)

                price_type = master_col[4] if not ka_range else price_type[-1]

            in_range_flag = False
            dp_list = []

            for i in range(len(search_data)):
                start_date = datetime.strptime(search_data[master_col[7]][i], "%Y%m%d")
                end_date = datetime.strptime(search_data[master_col[8]][i], "%Y%m%d")

                if start_date <= data_date <= end_date:
                    in_range_flag = True
                    dp = search_data[price_type][i]

                    if ((price_type == master_col[5]) or (price_type == master_col[6])) & (not pd.notna(dp)):
                        msg = f"{dealer_id} 中 {product_id} 之 {price_type} 值為空，系統將使用 {master_col[4]} 的值。"
                        WSysLog("2", "SearchDp", msg)
                        ptype_error[row] = f"{price_type} 值為空。"
                        price_type = master_col[4]
                        dp = search_data[price_type][i]

                    dp_list.append(dp)

            if not in_range_flag:
                msg = f"此貨號未搜尋到起迄區間符合 {datetime.strftime(data_date, '%Y/%m/%d')} 之資料。"
                error_row[row] = msg

            else:
                dp = dp_list[-1]
                output.append(dp)

        msg = f"ErrorRow結果： {error_row} 。"
        WSysLog("1", "SearchDp", msg)
        return output, error_row, ptype_error

# 使用 product id 在 MasterFile 中找到對應的 std 價
def search_cost(input_data, dealer_id):
    print("Function: search_cost")
    result, master_data, _ = read_master_file()
    output, error_row = [], {}
    if result:
        master_col = master_data.columns.values
        for row in range(len(input_data)):
            product_id = str(input_data["Product ID"][row])
            data_date = input_data["Transaction Date"][row]
            search_data = master_data[(master_data[master_col[0]] == dealer_id) & \
                                        (master_data[master_col[1]] == product_id)]
            search_data = search_data.reset_index(drop=True)
            in_range_flag = False
            cost_list = []

            for i in range(len(search_data)):
                start_date = datetime.strptime(search_data[master_col[7]][i], "%Y%m%d")
                end_date = datetime.strptime(search_data[master_col[8]][i], "%Y%m%d")

                if start_date <= data_date <= end_date:
                    in_range_flag = True
                    cost = search_data[master_col[3]][i]
                    cost_list.append(cost)

            if not in_range_flag:
                msg = f"此貨號未搜尋到起迄區間符合 {datetime.strftime(data_date, '%Y/%m/%d')} 之資料。"
                error_row[row] = msg

            else:
                cost = cost_list[-1]
                output.append(cost)

        msg = f"ErrorRow結果： {error_row} 。"
        WSysLog("1", "SearchCost", msg)
        return output, error_row

# 多欄位值合併
def merge_columns(input_data, source_col, value):
    parts = source_col.split("+")
    output = input_data[parts].apply\
            (lambda row: value.join\
            (row.values.astype(str)), axis=1)
    return output

# 取得月份起始與結束值
def getMonthStartAndEndDate(file_name):
    file_name_part = re.split(r"[._]" ,file_name)
    file_name_date = file_name_part[2]
    end_date, date_format = 0, False
    file_name_date_list = list(str(file_name_date))

    # 檔名中的日期時間只接受8碼或12碼
    if (len(file_name_date_list) == 8) or (len(file_name_date_list) == 12):
        date_format = True

    if not date_format:
        raise ValueError("檔名中的時間日期內容不符合規範。")

    else:
        year_list = file_name_date_list[0:4]
        year = int("".join(year_list))
        month_list = file_name_date_list[4:6]
        month = int("".join(month_list))

        if month < 8:
            if month == 2:
                end_date = 28
            if ((month - 1) % 2) == 0:
                end_date = 31
            else:
                end_date = 30

        else:
            if (month % 2) == 0:
                end_date = 31
            else:
                end_date = 30

        start_date = str(year) + str(month) + "01"
        end_date = str(year) + str(month) + str(end_date)
        return start_date, end_date

# 依據轉換規則轉換銷售檔案
def ChangeSaleFile(dealer_id, file_name):
    print("Function:ChangeSaleFile")
    print(f"dealer_id:{dealer_id}")
    print(f"file_name:{file_name}")
    change_status = "OK"
    file_header = Config.SaleOutputFileHeader
    change_rules = Config.SaleFileChangeRule

    for i in range(len(Config.DealerList)):
        if Config.DealerList[i] == dealer_id:
            index = i + 1
            break

    dealer_name = Config.DealerConfig[f"Dealer{index}"]["DealerName"]
    dealer_country = Config.DealerConfig[f"Dealer{index}"]["DealerCountry"]
    dealer_OUP_type = Config.DealerConfig[f"Dealer{index}"]["SaleFile"]["OUPType"]
    file_path = os.path.join(Config.DealerFolderPath, dealer_id, file_name)
    input_data = read_data(file_path)
    output_data = pd.DataFrame(columns = file_header)
    error_data, error_index = check_product_id(dealer_id, input_data)
    if len(error_data) > 0:
        change_status = "NO"
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
        if (target_col == "Gross Revenue") and (dealer_OUP_type == False):
            rule = "SearchDP"
        if (target_col == "Demand Class Desc") and (dealer_OUP_type == False):
            rule = "SearchCost"

        # 欄位為固定值
        if rule == "FixedValue":
            if target_col == "Area":
                value = dealer_country

            elif target_col == "Branch":
                value = dealer_id
            output_data[target_col] = fixed_value(value, len(input_data))

        # 搬移資料
        elif rule == "Move":
            output_data[target_col] = move_rule(input_data, source_col)

        # 變更時間格式
        elif rule == "ChangeTimeFormat":
            output_data[target_col] = change_time_format(input_data, source_col, value)

        # Quantity特殊處理
        elif rule == "MoveOrSearchUom":
            # target_col = "Original Quantity"
            output, error_row = move_or_search_uom(input_data, source_col, target_col, dealer_id)

            for row, msg in reversed(list(error_row.items())):
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                input_data = input_data.drop(row)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
            input_data = input_data.reset_index(drop = True)

        # 搜索MasterFile的dp價
        elif rule == "SearchDP":
            output, error_row, ptype_error = search_dp(input_data, dealer_id)

            for row, msg in ptype_error.items():
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error

            for row, msg in reversed(list(error_row.items())):
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                input_data = input_data.drop(row)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
            input_data = input_data.reset_index(drop = True)

        elif rule == "SearchCost":
            output, error_row = search_cost(input_data, dealer_id)

            for row, msg in reversed(list(error_row.items())):
                change_status = "NO"
                output_data = output_data.drop(row).reset_index(drop=True)
                input_data = input_data.drop(row)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
            input_data = input_data.reset_index(drop = True)

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
    start_date, end_date = getMonthStartAndEndDate(file_name)
    changed_file_name = Config.SaleOutputFileName.replace("{DealerID}", str(dealer_id))\
        .replace("{StartDate}", start_date)\
        .replace("{EndDate}", end_date)

    try:
        output_data.to_csv(os.path.join(Config.ChangeFolderPath,\
            f"{changed_file_name}.{Config.SaleOutputFileExtension}"), index=False)
        msg = f"檔案轉換完成，輸出檔名 {changed_file_name}.{Config.SaleOutputFileExtension}。"
        WChaLog("1", "ChangeSaleFile", dealer_id, file_name, msg)

        # error report
        error_file_name, error_report_path = None, None
        if not error_data.empty:
            error_file_name = Config.SaleErrorReportFileName.replace("{DealerID}", str(dealer_id))\
                .replace("{StartDate}", start_date)\
                .replace("{EndDate}", end_date)
            error_report_folder_path = os.path.join(Config.ErrorReportPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))

            if not os.path.exists(error_report_folder_path):
                os.makedirs(error_report_folder_path)
                msg = f"成功建立目錄 {error_report_folder_path} 。"
                WSysLog("1", "ChangeSaleFile", msg)

            error_report_path = os.path.join(error_report_folder_path, error_file_name)
            error_data.to_excel(error_report_path, index = False)
            msg = f"Error檔案輸出完成，輸出檔名 {error_file_name}。"
            WChaLog("1", "ChangeSaleFile", dealer_id, file_name, msg)
        chang_result = {"Status" : change_status, "OutputFileName": f"{changed_file_name}.{Config.SaleOutputFileExtension}",\
                        "ErrorNum": len(error_data), "Num" : len(output_data), "ErrorReportFileName" : error_file_name, "ErrorReportPath" : error_report_path}
        return  chang_result

    except Exception as e:
        change_status = "NO"
        msg = f"轉換檔案失敗，遇到未知錯誤，{e}。"
        WChaLog("2", "ChangeSaleFile", dealer_id, file_name, msg)
        chang_result = {"Status" : change_status, "OutputFileName": None,\
                        "ErrorNum": 0, "Num" : 0, "ErrorReportFileName" : None, "ErrorReportPath" : None}
        return chang_result

# 新欄位內容填寫 Transaction Date 最後一天
def last_transaction_date(input_data, row):
    return [str(datetime.strftime(input_data["Transaction Date"][row - 1], "%m/%d/%Y"))] * row

# 依據規則轉換庫存檔案
def ChangeInventoryFile(dealer_id, file_name):
    print("Function: ChangeInventoryFile")
    print(f"daeler_id:{dealer_id}")
    print(f"file_name:{file_name}")
    change_status = "OK"
    file_header = Config.InventoryOutputFileHeader
    change_rules = Config.InventoryFileChangeRule

    for i in range(len(Config.DealerList)):
        if Config.DealerList[i] == dealer_id:
            index = i + 1
            break

    dealer_name = Config.DealerConfig[f"Dealer{index}"]["DealerName"]
    file_path = os.path.join(Config.DealerFolderPath, dealer_id, file_name)
    input_data = read_data(file_path)
    output_data = pd.DataFrame(columns = file_header)
    error_data, error_index = check_product_id(dealer_id, input_data)

    if len(error_data) > 0:
        change_status = "NO"
    msg = f"檔案中有 {len(error_data)} 筆資料在 master file 貨號中找不到。"
    WChaLog("2","ChangeInventoryFile", dealer_id, file_name, msg)

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
                input_data = input_data.drop(row)
                new_error = input_data.iloc[row].to_dict()
                new_error["Dealer ID"] = dealer_id
                new_error["Dealer Name"] = dealer_name
                new_error["Exchange Error Issue"] = msg
                error_data.loc[len(error_data)] = new_error
            output_data[target_col] = output
            input_data = input_data.reset_index(drop = True)

        # Date Period欄位為 Transaction Date 最後一天
        elif rule == "LastTransactionDate":
            output_data[target_col] = last_transaction_date(input_data, len(input_data))

    start_date, end_date = getMonthStartAndEndDate(file_name)
    data_end_date = datetime.strftime(input_data["Transaction Date"][len(input_data) - 1], "%Y%m%d")
    changed_file_name = f"{dealer_id}_I_{data_end_date}.csv"

    try:
        output_data.to_csv(os.path.join(Config.ChangeFolderPath, changed_file_name), index=False)
        msg = f"檔案轉換完成，輸出檔名 {changed_file_name}。"
        WChaLog("1", "ChangeInventoryFile", dealer_id, file_name, msg)

        # error report
        error_file_name, error_report_path = None, None
        if not error_data.empty:
            error_file_name = Config.InventoryErrorReportFileName.replace("{DealerID}", str(dealer_id))\
                .replace("{StartDate}", start_date)\
                .replace("{EndDate}", end_date)
            error_report_folder_path = os.path.join(Config.ErrorReportPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))

            if not os.path.exists(error_report_folder_path):
                os.makedirs(error_report_folder_path)
                msg = f"成功建立目錄 {error_report_folder_path} 。"
                WSysLog("1", "ChangeSaleFile", msg)

            error_report_path = os.path.join(error_report_folder_path, error_file_name)
            error_data.to_excel(error_report_path, index = False)
            msg = f"Error檔案輸出完成，輸出檔名 {error_file_name}。"
            WChaLog("1", "ChangeInventoryFile", dealer_id, file_name, msg)
        change_result = {"Status" : change_status, "OutputFileName": changed_file_name, "ErrorNum": len(error_data),\
                         "Num" : len(output_data), "ErrorReportFileName" : error_file_name, "ErrorReportPath" : error_report_path}
        return change_result

    except Exception as e:
        change_status = "NO"
        msg = f"轉換檔案失敗，遇到未知錯誤，{e}。"
        WChaLog("2", "ChangeInventoryFile", dealer_id, file_name, msg)
        change_result = {"Status" : change_status, "OutputFileName": None, "ErrorNum": 0,\
                         "Num" : 0, "ErrorReportFileName" : None, "ErrorReportPath" : None}
        return change_result

# 轉換主程式
def Changing(check_right_list):
    for dealer_id in Config.DealerList:
        dealer_path = os.path.join(Config.DealerFolderPath, dealer_id)
        file_names = [file for file in os.listdir(dealer_path)\
                      if os.path.isfile(os.path.join(dealer_path, file))]
        error_files, error_paths = [], []

        for file_name in file_names:
            file_type, _ = get_file_type(dealer_id, file_name)

            if file_type == "Sale":
                change_result = ChangeSaleFile(dealer_id, file_name)
                status = change_result["Status"]
                output_file_name = change_result["OutputFileName"]
                error_num = change_result["ErrorNum"]
                num = change_result["Num"]
                error_report = change_result["ErrorReportFileName"]
                error_report_path = change_result["ErrorReportPath"]

                if error_num != 0:
                    error_files.append(error_report)
                    error_paths.append(error_report_path)

            else:
                change_result = ChangeInventoryFile(dealer_id, file_name)
                status = change_result["Status"]
                output_file_name = change_result["OutputFileName"]
                error_num = change_result["ErrorNum"]
                num = change_result["Num"]
                error_report = change_result["ErrorReportFileName"]
                error_report_path = change_result["ErrorReportPath"]
                if error_num != 0:
                    error_files.append(error_report)
                    error_paths.append(error_report_path)

            for data_id, name in check_right_list.items():
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

        if error_files:
            mail_data = {"ErrorReportFileName" : "、".join(error_files)}
            send_info = {"Mode" : "ErrorReport", "DealerID" : dealer_id, "MailData" : mail_data, "FilesPath" : error_paths}
            SendMail(send_info)

# 合併 Inventory 檔案
def MergeInventoryFile():
    file_names = [file for file in os.listdir(Config.ChangeFolderPath) \
                      if os.path.isfile(os.path.join(Config.ChangeFolderPath, file))]

    # 取得要合併的檔案
    file_list, time_list = [], []
    for file_name in file_names:
        part = re.split(r"[._]" ,file_name)
        if part[1] == "I":
            file_list.append(file_name)
            time_list.append(part[2])

    if file_list:
        # 抓取檔名中的檔案時間
        month = int(Config.Month)
        if month < 8:
            if month == 2:
                end_date = 28
            if ((month - 1) % 2) == 0:
                end_date = 31
            else:
                end_date = 30
        else:
            if (month % 2) == 0:
                end_date = 31
            else:
                end_date = 30
        end_date = str(Config.Year) + str(month) + str(end_date)

        # 合併檔案
        dataframes = [pd.read_csv(os.path.join(Config.ChangeFolderPath, file)) for file in file_list]
        combined_df = pd.concat(dataframes, ignore_index=True)

        # 輸出檔案
        changed_file_name = Config.InventoryOutputFileName.replace("{CountryCode}", Config.InventoryOutputFileCountryCode)\
            .replace("{LastDate}", end_date)
        try:
            combined_df.to_csv(os.path.join(Config.ChangeFolderPath, \
                f"{changed_file_name}.{Config.InventoryOutputFileExtension}"), sep = ",", index=False)
            msg = f"{len(file_list)} 份檔案成功合併成 {changed_file_name}.{Config.InventoryOutputFileExtension}。"
            WSysLog("1", "MargeInventoryFile", msg)
            for dealer_id in Config.DealerList:
                target_folder = os.path.join(Config.ChangeFolderPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)

                for file in file_list:
                    part = re.split(r"[._]" ,file)
                    if (part[0] == dealer_id) and (part[1] == "I"):
                        file_source = os.path.join(Config.ChangeFolderPath, file)
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

# 檔案上傳EFT雲端完成後歸檔
def FileArchiving():
    file_names = [file for file in os.listdir(Config.ChangeFolderPath) \
                if os.path.isfile(os.path.join(Config.ChangeFolderPath, file))]
    
    for file in file_names:
        part = file.split("_")
        for dealer_id in Config.DealerList:
            if part[0] == dealer_id:
                target_folder = os.path.join(Config.ChangeFolderPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                file_source = os.path.join(Config.ChangeFolderPath, file)
                file_target = os.path.join(target_folder, file)
                shutil.move(file_source, file_target)
                if os.path.exists(file_target):
                    msg = f"檔案搬移至 {target_folder} 成功"
                    WSysLog("1", "MoveInventoryFile", msg)
                else:
                    msg = f"檔案搬移至 {target_folder} 失敗"
                    WSysLog("2", "MoveInventoryFile", msg)

        if part[0] == Config.InventoryOutputFileCountryCode:
            target_folder = os.path.join(Config.ChangeFolderPath, Config.MergeInventoryFolder, datetime.strftime(Config.SystemTime, "%Y%m"))
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)
            file_source = os.path.join(Config.ChangeFolderPath, file)
            file_target = os.path.join(target_folder, file)
            shutil.move(file_source, file_target)
            if os.path.exists(file_target):
                msg = f"檔案搬移至 {target_folder} 成功"
                WSysLog("1", "MoveInventoryFile", msg)
            else:
                msg = f"檔案搬移至 {target_folder} 失敗"
                WSysLog("2", "MoveInventoryFile", msg)

if __name__ == "__main__":
    # aa = {17: '111_I_20240726.csv', 18: '111_S_20240726.csv'}
    # Changing(aa)
    # ChangeSaleFile(dealerID, FilePath)
    # input_data, data_max_row = read_data(FilePath)
    # dealerID = "111"
    # FilePath = "111_I_20240726.csv"
    # ChangeInventoryFile(dealerID, FilePath)
    # check_product_id(dealerID, input_data)
    # MergeInventoryFile()
    FileArchiving()

# ===================================================================

'''
# # 標準庫
# import os, re, math, shutil
# from datetime import datetime
# from dateutil import parser

# # 第三方庫
# import pandas as pd

# # 自定義函數
# from Mail import SendMail
# from Log import WSysLog, WChaLog
# from CheckFile import get_file_type
# from RecordTable import WriteSubRawData
# from Config import AppConfig

# Config = AppConfig()

# # 程式運作的對應開關。設置Ture則為開啟，False則為關閉
# # 此開關用於顯示轉換過程中的相關運作進度
# ShowScheduleSwitch = False

# # 此開關用於顯示轉換過程中的資料處理
# ShowRunningDataSwitch = False

# class dataChangeFuncion:
#     """
#     MasterFile Header：
#         經銷商號碼\nSold-to code, 貨號,	UOM, std cost (EA),	DP(EA),	KADP(EA), IVY (EA),	起,	迄

#     KAList Header：
#         經銷商ID, 客戶號, 起, 迄, Price TYPE

#     轉換過程中，共用的函數與轉換規則對應函數

#         class 內部函數呼叫使用
#             函數名稱                     說明
#             __init__                    class 內區域共用變數
#             print_master_file_data      顯示 masterfile 的資料
#             print_kalist_file_data      顯示 kalist 的資料
#             parse_and_format_date       統一時間欄位內容格式；date_str -> str；output_format -> "%Y/%m/%d"

#             search_pid_in_master_file   在MasterFile中搜尋產品ID，回傳需要的 value 值；
#                                         dealer_id -> str；product_id -> str；data_date -> time format

#             get_dp_type_in_kalist       從 KAList 工作表中搜尋 price type；
#                                         dealer_id -> str；buyer_id -> str；data_date -> time format

#         轉換過程中共用函數
#             函數名稱                     說明
#             getFilePath                 取得檔案目錄；dealer_id -> str；file_name -> str
#             delFile                     刪除檔案；file_path -> str
#             getMonthStartAndEndDate     取得檔案名稱對應的起始日期與結束日期；file_name -> str
#             checkProductIdValue         確認檔案 product_id 欄位值非中文；file_path -> str

#             checkProdictIdInMasterFile  比對檔案中的 product id 存在於 master file 中；
#                                         input_file_data -> str；dealer_id -> str

#         轉換規則對應的函數
#             函數名稱                     說明
#             moveRule                    將原先欄位的值移動到搬移到新的欄位；input_data -> DataFrame；
#             fixedValue                  欄位值固定為某些值；value -> str；row -> int

#             changeTimeFormat            轉換欄位內容的時間格式；
#                                         input_data -> DataFrame；input_col -> str；data_format -> str

#             moveOrSearchUom             搬移或是轉換 Uom 值；
#                                         input_data -> DataFrame；source_col -> str；
#                                         dealer_id -> str；target_col -> str

#             searchDP                    搜尋 Dp 資料，若在 kalist 中，則需篩選對應的 value；
#                                         input_data -> DataFrame；dealer_id -> str

#             searchCost                  搜尋 Cost 資料；
#                                         input_data -> DataFrame；dealer_id -> str

#             mergeColumns                多欄位值合併；
#                                         input_data -> DataFrame；source_col -> str；value -> str

#             lastTransactionDate         將 transaction date 依照升序排序後，取出最後一個值；
#                                         input_data -> DataFrame；row -> int
#     """

#     # class 內區域變數
#     def __init__(self):
#         # class內通用參數
#         self.master_file_header = ["經銷商號碼\nSold-to code",
#                                    "貨號",
#                                    "UOM",
#                                    "std cost (EA)",
#                                    "DP(EA)",
#                                    "KADP(EA)",
#                                    "IVY (EA)",
#                                    "起",
#                                    "迄"]

#         self.kalist_file_header = [ "經銷商ID",
#                                     "客戶號",
#                                     "起",
#                                     "迄",
#                                     "Price TYPE"]

#         # 固定值
#         self.product_id = "Product ID"
#         self.transaction_date = "Transaction Date"
#         self.creation_date = "Creation Date"
#         self.buyer_id = "Buyer ID"
#         self.defult_price_type = "DP"

#         # 從Config檔案抓取參數
#         self.dealer_folder_path = Config.DealerFolderPath
#         self.error_report_folder_path = Config.ErrorReportPath

#         date_cols = ["起", "迄"]
#         master_folder_path = Config.MasterFolderPath
#         # 測試用
#         # master_folder_path = "./datas"
#         master_file_name = "MasterFile.xlsx"
#         master_file_sheet_name = "MasterFile"
#         kalist_file_sheet_name = "KAList"

#         master_file_path = os.path.join(master_folder_path, master_file_name)

#         try:
#             # MasterFile 資料
#             self.master_file_data = pd.read_excel(master_file_path,
#                                                 sheet_name = master_file_sheet_name,
#                                                 dtype = str)

#             # 比對 MasterFile 工作表要與規定的符合
#             master_file_data_header = self.master_file_data.columns.values
#             if set(self.master_file_header) != set(master_file_data_header):
#                 msg = "MasterFile 工作表的表頭與規定值不匹配。"
#                 raise ValueError(msg)

#             for col in date_cols:
#                 self.master_file_data[col] = pd.to_datetime(self.master_file_data[col],
#                                                             format = "%Y%m%d")

#             # KAList 資料
#             self.kalist_file_data = pd.read_excel(master_file_path,
#                                                 sheet_name = kalist_file_sheet_name,
#                                                 dtype = str)

#             # 比對 KAList 工作表要與規定的符合
#             kalist_file_data_header = self.kalist_file_data.columns.values
#             if set(self.kalist_file_header) != set(kalist_file_data_header):
#                 msg = "KAList 工作表的表頭與規定值不匹配。"
#                 raise ValueError(msg)

#             for col in date_cols:
#                 self.kalist_file_data[col] = pd.to_datetime(self.kalist_file_data[col],
#                                                             format = "%Y%m%d")

#             msg = "成功讀取 MasterFile 資料。"
#             WSysLog("1", "ReadMasterFile", msg)

#         except Exception as e:
#             msg = f"讀取 MasterFile 時發生錯誤。錯誤原因：{str(e)}"
#             WSysLog("3", "ReadMasterFile", msg)
#             raise FileNotFoundError(msg) from e

#     # 將 master_file 內容輸出
#     def print_master_file_data(self):
#         if ShowScheduleSwitch:
#             print("Function master_file_test start.")
#         print(self.master_file_data)
#         if ShowScheduleSwitch:
#             print("Function master_file_test finish.")

#     # 將 kalist 內容輸出
#     def print_kalist_file_data(self):
#         if ShowScheduleSwitch:
#             print("Function kalist_file_test start.")
#         print(self.kalist_file_data)
#         if ShowScheduleSwitch:
#             print("Function kalist_file_test finish.")

#     # 取得檔案目錄
#     def getFilePath(self, dealer_id, file_name):
#         if ShowScheduleSwitch:
#             print("\t\tFunction get_file_path start.")
#             print("\t\tFunction get_file_path end.")
#         return os.path.join(self.dealer_folder_path, dealer_id, file_name)

#     # 取得錯誤報表存放目錄
#     def getErrorReportPath(self, dealer_id):
#         if ShowScheduleSwitch:
#             print("\t\tFunction get_error_report_path start.")
#         folder_name = datetime.strftime(Config.SystemTime, "%Y%m")
#         if ShowScheduleSwitch:
#             print("\t\tFunction get_error_report_path end.")
#         return os.path.join(self.error_report_folder_path, dealer_id, folder_name)

#     # 刪除檔案
#     def delFile(self, file_path):
#         if ShowScheduleSwitch:
#             print("\tFunction del_file start.")

#         try:
#             os.remove(file_path)
#             if ShowScheduleSwitch:
#                 print("\tFunction del_file end.")
#         except Exception as e:
#             msg = f"系統刪除檔案時發生錯誤。錯誤原因:{str(e)}"
#             WSysLog("3", "DelFile", msg)
#             raise OSError(msg) from e

#     # 取得檔案名稱對應的起始日期與結束日期
#     def getMonthStartAndEndDate(self, file_name):
#         if ShowScheduleSwitch:
#             print("\t\tFunction get_month_start_and_end_date start.")
#         file_name_part = re.split(r"[._]" ,file_name)
#         file_name_date = file_name_part[2]
#         end_date, date_format = 0, False
#         file_name_date_list = list(str(file_name_date))

#         # 檔名中的日期時間只接受8碼或12碼
#         if (len(file_name_date_list) == 8) or (len(file_name_date_list) == 12):
#             date_format = True

#         if not date_format:
#             raise ValueError("檔名中的時間日期內容不符合規範。")

#         else:
#             year_list = file_name_date_list[0:4]
#             year = int("".join(year_list))
#             month_list = file_name_date_list[4:6]
#             month = int("".join(month_list))

#             if month < 8:
#                 if month == 2:
#                     end_date = 28
#                 if ((month - 1) % 2) == 0:
#                     end_date = 31
#                 else:
#                     end_date = 30

#             else:
#                 if (month % 2) == 0:
#                     end_date = 31
#                 else:
#                     end_date = 30

#             start_date = str(year) + str(month) + "01"
#             end_date = str(year) + str(month) + str(end_date)
#             if ShowScheduleSwitch:
#                 print("\t\tFunction get_month_start_and_end_date end.")
#             return start_date, end_date

#     # 統一時間欄位內容格式
#     def parse_and_format_date(self, date_str, output_format = "%Y/%m/%d"):
#         # print("Function parse_and_format_date start.")
#         try:
#             parsed_date = parser.parse(date_str)
#             return parsed_date.strftime(output_format)
#         except (ValueError, TypeError):
#             return date_str
#         # finally:
#         #     print("Function parse_and_format_date finish.")

#     # 確認檔案 product_id 欄位值非中文
#     def checkProductIdValue(self, file_path):
#         if ShowScheduleSwitch:
#             print("\t\tFunction check_product_id_value start.")
#         file_name = os.path.basename(file_path)
#         _, file_extension = os.path.splitext(file_name)
#         file_extension = file_extension.lower()

#         if file_extension in Config.AllowFileExtensions:
#             file_data = pd.read_csv(file_path, dtype = str)\
#                 if file_extension == ".csv"\
#                 else pd.read_excel(file_path, dtype = str)

#             # print(file_data)

#             # 將時間欄位內容資料型態統一
#             date_cols = [self.transaction_date, self.creation_date]
#             if ShowScheduleSwitch:
#                 print("\t\tFunction parse_and_format_date start.")
#             for col in date_cols:
#                 file_data[col] = file_data[col].apply(lambda x: self.parse_and_format_date(str(x)))
#                 file_data[col] = pd.to_datetime(file_data[col],
#                                                 format = "%Y/%m/%d",
#                                                 errors = "coerce")
#             file_data[self.product_id] = file_data[self.product_id].astype(str).str.strip()
#             if ShowScheduleSwitch:
#                 print("\t\tFunction parse_and_format_date finish.")

#             # 過濾 Product ID 僅允許 a-z, A-Z, 0-9, - 符號
#             file_data = file_data[file_data[self.product_id].str.contains\
#                 ("^[a-zA-Z0-9-]+$", regex=True, na=False)]

#             # 刷新index值
#             file_data = file_data.reset_index()
#             if ShowScheduleSwitch:
#                 print("\t\tFunction check_product_id_value end.")
#             return file_data

#         else:
#             msg = f"檔案附檔名不符合規範。檔名：{file_name}"
#             WSysLog("2", "CheckProductIdValue", msg)
#             if ShowScheduleSwitch:
#                 print("\t\tFunction check_product_id_value end.")
#             return False

#     # 比對檔案中的 product id 存在於 master file 中
#     def checkProdictIdInMasterFile(self, input_file_data, dealer_id):
#         if ShowScheduleSwitch:
#             print("\t\tFunction check_prodict_id_in_master_file start.")
#         col_dealer_id = self.master_file_header[0]
#         col_product_id = self.master_file_header[1]
#         pid_not_in_master_file = []
#         # input_file_data -> dataFrame
#         # 取出輸入資料中的 "product id" 欄位資料，去掉重複值，排序
#         input_data_product_id = input_file_data.loc[:, self.product_id].tolist()
#         input_data_product_id = sorted(list(set(input_data_product_id)))
#         # print(input_data_product_id)
#         # print(len(input_data_product_id))

#         # master_file_data -> dataFrame，此處輸入的masterfile是該經銷商的資料，非整份masterfile
#         # 取出輸入資料中的 "貨號" 欄位資料，去掉重複值，排序
#         master_file_data = self.master_file_data[\
#             self.master_file_data[col_dealer_id] == dealer_id]
#         master_file_product_id = master_file_data.loc[:, col_product_id].tolist()
#         master_file_product_id = sorted(list(set(master_file_product_id)))

#         # 使用迴圈檢查經銷商上傳檔案中的 product_id 是否存在於 master_file 資料中
#         for pid in input_data_product_id:
#             if pid not in master_file_product_id:
#                 pid_not_in_master_file.append(pid)
#         # print(f"pid_not_in_master_file:{pid_not_in_master_file}")

#         input_data_in_master_file = input_file_data[\
#             ~input_file_data[self.product_id].isin(pid_not_in_master_file)]
#         # print("input_data_in_master_file")
#         # print(input_data_in_master_file)

#         input_data_not_in_master_file = input_file_data[\
#             input_file_data[self.product_id].isin(pid_not_in_master_file)]
#         # print("input_data_not_in_master_file")
#         # print(input_data_not_in_master_file)
#         if ShowScheduleSwitch:
#             print("\t\tFunction check_prodict_id_in_master_file end.")
#         return input_data_in_master_file, input_data_not_in_master_file

#     # 將原先欄位的值移動到搬移到新的欄位
#     def moveRule(self, input_data, input_col):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction move_rule start.")
#             print("\t\t\tFunction move_rule end.")
#         # print(input_data[input_col])
#         return input_data[input_col].to_list()

#     # 欄位值固定為某些值
#     def fixedValue(self, value, row):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction fixed_value start.")
#             print("\t\t\tFunction fixed_value end.")
#         return [value] * row

#     # 轉換欄位內容的時間格式
#     def changeTimeFormat(self, input_data, input_col, date_format):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction change_time_format start.")
#             print("\t\t\tFunction change_time_format end.")
#         # print(input_data[input_col].dt.strftime(date_format))
#         return input_data[input_col].dt.strftime(date_format).to_list()

#     # 在MasterFile中搜尋產品ID，回傳需要的 value 值
#     def search_pid_in_master_file(self, dealer_id, product_id, data_date):
#         # print("Function search_pid_in_master_file start.")
#         col_dealer_id = self.master_file_header[0]
#         col_product_id = self.master_file_header[1]
#         col_start_date = self.master_file_header[7]
#         col_end_date = self.master_file_header[8]

#         # 從 master_file 中取出對應 經銷商ID 的資料
#         master_file_data = self.master_file_data\
#             [self.master_file_data[col_dealer_id] == dealer_id]

#         search_pid_data = master_file_data[master_file_data[col_product_id] == product_id]

#         if not search_pid_data.empty:
#             # 透過data_date篩選除對應區間的masterfile資料
#             pid_data_in_date = search_pid_data[(search_pid_data[col_start_date] <= data_date) &
#                                                 (search_pid_data[col_end_date] >= data_date)]
#             if not pid_data_in_date.empty:
#                 # print("Function search_pid_in_master_file end.")
#                 return True, pid_data_in_date

#             else:
#                 msg = f"經銷商：{dealer_id} 的產品ID：{product_id}，在 masterfile 工作表中搜尋不到對應的起迄區間。"
#                 WSysLog("3", "SearchPidInMasterFile", msg)
#                 msg = "在 masterfile 檔案中搜尋不到對應的起迄區間。"
#                 # print("Function search_pid_in_master_file end.")
#                 return None, msg
#         else:
#             msg = f"經銷商：{dealer_id} 的產品ID：{product_id}，在 masterfile 工作表中搜尋不到。"
#             WSysLog("3", "SearchPidInMasterFile", msg)
#             msg = "在 masterfile 檔案中搜尋不到此 Product ID。"
#             # print("Function search_pid_in_master_file end.")
#             return False, msg

#     # 搬移或是轉換 Uom 值
#     # 產出的錯誤列表需先從原始資料移除對應行，否則會對應不上
#     def moveOrSearchUom(self, input_data, source_col, dealer_id, target_col):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction move_or_search_uom start.")
#         col_uom = self.master_file_header[2]
#         no_value_in_masterfile = {}

#         value_list_in_source = input_data[source_col].tolist()
#         # print(value_list_in_source)

#         # 篩選來源col欄位中值為空白的row，並取得index
#         source_na_index_list = input_data[input_data[source_col].isna()].index.tolist()

#         for row in source_na_index_list:
#             # print(row)
#             input_data_product_id = input_data.loc[row, self.product_id]
#             input_data_date = input_data.loc[row, self.transaction_date]
#             # print(f"input_data_product_id:{input_data_product_id}")
#             # print(f"input_data_date:{input_data_date}")
#             search_result, pid_data_in_date = self.search_pid_in_master_file\
#                 (dealer_id, input_data_product_id, input_data_date)

#             if search_result:
#                 # 從搜尋結果中取出全部的UOM值
#                 target_col_list = pid_data_in_date[col_uom].to_list()
#                 try:
#                     # 取出最後的值
#                     value = target_col_list[-1]
#                     if (isinstance(value, float)) and (math.isnan(value)):
#                         row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
#                         msg = f"masterfile檔案中 {col_uom} 欄位第 {row_in_masterfile} 行數值為空。"
#                         WSysLog("2", "MoveOrSearchUom", msg)
#                         msg = f"{col_uom} 欄位第 {row_in_masterfile} 行數值為空。"
#                         no_value_in_masterfile[row] = msg

#                     elif isinstance(value, str):
#                         uom_in_masterfile = int(value)
#                         # print(input_data.loc[row, target_col])
#                         # print(type(input_data.loc[row, target_col]))
#                         if (isinstance(input_data.loc[row, target_col], float)) and\
#                             (math.isnan(input_data.loc[row, target_col])):
#                             msg = f"經銷商寫入的 {target_col} 欄位於第 {row + 2} 數值為空。"
#                             WSysLog("2", "MoveOrSearchUom", msg)
#                             raise ValueError(msg)

#                         else:
#                             changed_value = uom_in_masterfile *\
#                                 float(input_data.loc[row, target_col])
#                             value_list_in_source[row] = str(changed_value)

#                 except Exception as e:
#                     msg = f"將搜尋的資訊轉換為數值時發生錯誤。錯誤原因：{str(e)}"
#                     WSysLog("3", "MoveOrSearchUom", msg)
#                     raise TypeError(msg) from e

#             else:
#                 no_value_in_masterfile[row] = pid_data_in_date

#         # print("no_value_in_masterfile")
#         # print(no_value_in_masterfile)
#         # 移除掉無法轉換的行
#         # for index in sorted(no_value_in_masterfile, reverse=True):
#             # print(f"index:{index}")
#             # del value_list_in_source[index]
#         # print(value_list_in_source)
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction move_or_search_uom end.")
#         return value_list_in_source, no_value_in_masterfile

#     # 從 KAList 工作表中搜尋 price type
#     # searchDP 專用
#     def get_dp_type_in_kalist(self, dealer_id, buyer_id, data_date):
#         # print("Function get_dp_type start.")
#         col_dealer_id_in_ka = self.kalist_file_header[0]
#         col_buyer_id_in_ka = self.kalist_file_header[1]
#         col_start_date = self.kalist_file_header[2]
#         col_end_date = self.kalist_file_header[3]
#         col_price_type = self.kalist_file_header[4]
#         price_type = self.defult_price_type

#         # 從 kalist 工作表中篩選經銷商與客戶號資訊
#         search_data_in_ka = self.kalist_file_data\
#             [(self.kalist_file_data[col_dealer_id_in_ka] == dealer_id) &
#             (self.kalist_file_data[col_buyer_id_in_ka] == buyer_id)]

#         if not search_data_in_ka.empty:
#             buyer_data_in_date = search_data_in_ka\
#                 [(search_data_in_ka[col_start_date] <= data_date) &
#                 (search_data_in_ka[col_end_date] >= data_date)]

#             if not buyer_data_in_date.empty:
#                 type_list = buyer_data_in_date[col_price_type].to_list()
#                 price_type = type_list[-1]
#                 # print("Function get_dp_type end.")
#                 return price_type
#             else:
#                 # msg = f"經銷商ID：{dealer_id} 的客戶號：{buyer_id} 資料，在 KAList 工作表中未搜尋到符合時間區間的資料。"
#                 # WSysLog("2", "get_dp_type_in_kalist", msg)
#                 # print("Function get_dp_type end.")
#                 return price_type
#         else:
#             # msg = f"經銷商ID：{dealer_id} 的客戶號：{buyer_id} 資料，在 KAList 工作表中搜尋不到 。"
#             # WSysLog("2", "get_dp_type_in_kalist", msg)
#             # print("Function get_dp_type end.")
#             return price_type

#     # 搜尋 Dp 資料，若在 kalist 中，則需篩選對應的 value
#     # 產出的錯誤列表需先從原始資料移除對應行，否則會對應不上
#     def searchDP(self, input_data, dealer_id):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction search_dp start.")
#         dp_output, no_value_in_masterfile = [], {}
#         col_dp = self.master_file_header[4]

#         # for row in range(len(input_data)):
#         for row_index, single_row in input_data.iterrows():
#             # print(row)
#             # print(type(row))
#             input_data_buyer_id = single_row[self.buyer_id]
#             input_data_prodict_id = single_row[self.product_id]
#             input_data_date = single_row[self.transaction_date]

#             # 搜尋客戶號price_type
#             price_type = self.get_dp_type_in_kalist(dealer_id, input_data_buyer_id, input_data_date)
#             col_price_type = {"DP":"DP(EA)",
#                             "KA":"KADP(EA)",
#                             "IVY":"IVY (EA)"}.get(price_type, "DP(EA)")

#             # 在 masterfile 中搜尋 pid 對應區間的資料
#             search_result, pid_data_in_date = self.search_pid_in_master_file\
#                 (dealer_id, input_data_prodict_id, input_data_date)

#             if search_result:
#                 price_list = pid_data_in_date[col_price_type].to_list()
#                 dp_price_list = pid_data_in_date[col_dp].to_list()
#                 # print(f"price_list:{price_list}")
#                 # print(f"dp_price_list:{dp_price_list}")

#                 if len(price_list) == len(dp_price_list):
#                     price_result = price_list[-1]

#                     try:
#                         price_result = round(float(price_result), 2)
#                         dp_output.append(price_result)

#                     except (ValueError, TypeError) as e:
#                         if (isinstance(price_result, float)) and (math.isnan(price_result)):
#                             row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
#                             msg = f"masterfile檔案中 '{col_price_type}' 欄位第 {row_in_masterfile} 行數值為空。"
#                             WSysLog("2", "SearchDP", msg)
#                             msg = f"'{col_price_type}' 欄位第 {row_in_masterfile} 行數值為空。"
#                             no_value_in_masterfile[row_index] = msg
#                             dp_output.append(None)

#                         price_result = dp_price_list[-1]

#                         if (isinstance(price_result, float)) and (math.isnan(price_result)):
#                             row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
#                             msg = f"masterfile檔案中 'DP(EA)' 欄位第 {row_in_masterfile} 行數值為空。"
#                             WSysLog("2", "SearchDP", msg)
#                             msg = f"'DP(EA)' 欄位第 {row_in_masterfile} 行數值為空。"
#                             no_value_in_masterfile[row_index] = msg
#                             dp_output.append(None)

#                         elif isinstance(price_result, str):
#                             try:
#                                 # 小數點第三位，四捨五入進第二位
#                                 price_in_masterfile = round(float(price_result), 2)
#                                 dp_output.append(price_in_masterfile)

#                             except ValueError as e:
#                                 row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
#                                 msg = f"masterfile檔案中 'DP(EA)' 欄位第 {row_in_masterfile} 行轉換時發生錯誤。錯誤原因:{str(e)}"
#                                 WSysLog("2", "SearchDP", msg)
#                                 raise ValueError(msg) from e

#                 else:
#                     msg = "系統搜尋出的price_list與dp_price_list數量不匹配。"
#                     raise ValueError(msg)

#             else:
#                 # 比對masterfile後若無資料先新增空白值，後續統一移除
#                 dp_output.append(None)
#                 no_value_in_masterfile[row_index] = pid_data_in_date

#             # 測試用
#             # if row_index == 10:
#             #     break

#         # 測試用
#         # print("dp_output")
#         # print(dp_output)
#         # print(len(dp_output))
#         # print()
#         # print("no_value_in_masterfile")
#         # print(no_value_in_masterfile)
#         # print()
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction search_dp end.")
#         return dp_output, no_value_in_masterfile

#     # 搜尋 Cost 資料
#     def searchCost(self, input_data, dealer_id):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction search_cost start.")
#         cost_output, no_value_in_masterfile = [],{}
#         col_cost = self.master_file_header[3]

#         for row_index, single_row in input_data.iterrows():
#             # print(f"row_index:{row_index}")
#             input_data_prodict_id = single_row[self.product_id]
#             input_data_date = single_row[self.transaction_date]

#             # 在 masterfile 中搜尋 pid 對應區間的資料
#             search_result, pid_data_in_date = self.search_pid_in_master_file\
#                 (dealer_id, input_data_prodict_id, input_data_date)

#             if search_result:
#                 # print("pid_data_in_date")
#                 # print(pid_data_in_date)
#                 # print()
#                 cost_list = pid_data_in_date[col_cost].to_list()
#                 cost_value = cost_list[-1]

#                 try:
#                     cost_value = round(float(cost_value), 2)
#                     cost_output.append(cost_value)

#                 except (ValueError, TypeError) as e:
#                     if (isinstance(cost_value, float)) and (math.isnan(cost_value)):
#                         row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
#                         msg = f"masterfile檔案中 '{col_cost}' 欄位第 {row_in_masterfile} 行數值為空。"
#                         WSysLog("2", "SearchCost", msg)
#                         msg = f"'{col_cost}' 欄位第 {row_in_masterfile} 行數值為空。"
#                         no_value_in_masterfile[row_index] = msg
#                         cost_output.append(None)

#             else:
#                 # 比對masterfile後若無資料先新增空白值，後續統一移除
#                 cost_output.append(None)
#                 no_value_in_masterfile[row_index] = pid_data_in_date

#             # 測試用
#             # if row_index == 18:
#             #     break

#         # 測試用
#         # print("cost_output")
#         # print(cost_output)
#         # print(len(cost_output))
#         # print()
#         # print("no_value_in_masterfile")
#         # print(no_value_in_masterfile)
#         # print()
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction search_cost end.")
#         return cost_output, no_value_in_masterfile

#     # 多欄位值合併
#     def mergeColumns(self, input_data, source_col, value):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction merge_columns start.")
#         parts = source_col.split("+")
#         merge_output = input_data[parts].apply\
#                 (lambda row: value.join(row.values.astype(str)), axis=1)
#         # print(merge_output)
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction merge_columns end.")
#         return merge_output.to_list()

#     # 將 transaction date 依照升序排序後，取出最後一個值
#     def lastTransactionDate(self, input_data, row):
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction last_transaction_date start.")
#         transaction_list = input_data.loc[:, self.transaction_date].to_list()
#         # print("transaction_list")
#         # print(transaction_list)
#         transaction_list.sort()
#         last_value = transaction_list[-1].strftime("%m/%d/%Y")
#         # print(f"last_value:{last_value}")
#         if ShowScheduleSwitch:
#             print("\t\t\tFunction last_transaction_date end.")
#         return [last_value] * row

# class SaleDataChange(dataChangeFuncion):
#     """
#     銷售檔案轉換流程

#         class 內部函數呼叫使用
#             函數名稱                     說明
#             __init__                    class 內區域共用變數
#             convert_data_in_rule        根據sale轉換規則轉換來源資料
#             convert_output_info         產出 sale 轉換結果
#             convert_error_output_info   產出 sale 的 error report

#         外部呼叫函數
#             函數名稱                     說明
#             changeSaleFile              銷售檔案轉換主流程；dealer_id -> str；file_name -> str
#     """

#     # class 內區域變數
#     def __init__ (self):
#         super().__init__()

#         # 轉換規則
#         self.change_rule = Config.SaleFileChangeRule

#         # 轉換檔案產出位置
#         self.changed_folder_path = Config.ChangeFolderPath
#         # 測試用
#         # self.changed_folder_path = "./datas"

#         # 銷售轉換產出檔案副檔名
#         self.changed_ex = Config.SaleOutputFileExtension

#         # 設定預設值
#         self.dealer_id = None
#         self.dealer_name = None
#         self.file_name = None
#         self.dealer_country = None
#         self.dealer_OUP_type = None
#         self.input_file_data = None
#         self.in_master_file_data = None
#         self.not_in_master_file_data = None
#         self.convert_error = {}
#         self.change_status = "OK"
#         self.convert_output = None
#         self.change_result = {"Status" : self.change_status,
#                             "OutputFileName": None,
#                             "ErrorNum": 0,
#                             "Num" : 0,
#                             "ErrorReportFileName" : None,
#                             "ErrorReportPath" : None}

#     # 根據sale轉換規則轉換來源資料
#     def convert_data_in_rule(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_data_in_rule start.")
#         input_data = self.in_master_file_data

#         for column_index in self.change_rule:
#             # 從 mapping.json 取出銷售檔案每個 col 對應的轉換參數
#             source_col = self.change_rule[column_index]["SourceName"]
#             target_col = self.change_rule[column_index]["ColumnName"]
#             rule = self.change_rule[column_index]["ChangeRule"]
#             value = self.change_rule[column_index]["Value"]
#             if ShowRunningDataSwitch:
#                 print(f"rule:{rule}")

#             # 強制轉換規則
#             if (target_col == "Gross Revenue") and (not self.dealer_OUP_type):
#                 rule = "SearchDP"

#             if (target_col == "Demand Class Desc") and (not self.dealer_OUP_type):
#                 rule = "SearchCost"

#             # 欄位為固定值
#             if rule == "FixedValue":
#                 if target_col == "Area":
#                     value = self.dealer_country

#                 elif target_col == "Branch":
#                     value = self.dealer_id
#                 self.convert_output[target_col] = self.fixedValue\
#                     (value, len(input_data))

#             # 搬移資料
#             elif rule == "Move":
#                 self.convert_output[target_col] = self.moveRule\
#                     (input_data, source_col)

#             # 變更時間格式
#             elif rule == "ChangeTimeFormat":
#                 self.convert_output[target_col] = self.changeTimeFormat\
#                     (input_data, source_col, value)

#             # Quantity 特殊處理
#             elif rule == "MoveOrSearchUom":
#                 self.convert_output[target_col], no_value_in_masterfile = \
#                     self.moveOrSearchUom(input_data, source_col, self.dealer_id, target_col)
#                 self.convert_error.update(no_value_in_masterfile)

#             # 搜索 MasterFile 的 DP 價
#             elif rule == "SearchDP":
#                 self.convert_output[target_col], no_value_in_masterfile = \
#                     self.searchDP(input_data, self.dealer_id)
#                 self.convert_error.update(no_value_in_masterfile)

#             # 搜索 MasterFile 的 Cost
#             elif rule == "SearchCost":
#                 self.convert_output[target_col], no_value_in_masterfile = \
#                     self.searchCost(input_data, self.dealer_id)
#                 self.convert_error.update(no_value_in_masterfile)

#             # 多欄位內容合併
#             elif rule == "MergeColumns":
#                 self.convert_output[target_col] = self.mergeColumns\
#                     (input_data, source_col, value)

#             # 跳過不處理
#             elif not rule:
#                 continue

#             else:
#                 self.change_status = "NO"
#                 msg = f"{rule} 此轉換規則不再範圍中。"
#                 WSysLog("3", "ConvertDataInRule", msg)
#                 raise ValueError(msg)

#             if ShowRunningDataSwitch:
#                 print(self.convert_output)

#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_data_in_rule end.")

#     # 產出 sale 轉換結果
#     def convert_output_info(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_output_info start.")
#         # 取得經銷商上傳檔案之檔名對應的起始日期與結束日期
#         start_date, end_date = self.getMonthStartAndEndDate(self.file_name)
#         self.convert_output.reset_index(drop = True)

#         # 取得輸出的檔案名稱
#         convert_file_name = Config.SaleOutputFileName.replace("{DealerID}", str(self.dealer_id))\
#             .replace("{StartDate}", start_date).replace("{EndDate}", end_date)

#         # 加上副檔名，組成完整的檔案名稱
#         convert_file = f"{convert_file_name}.{self.changed_ex}"

#         # 資訊寫入轉換結果 dict 中
#         self.change_result["OutputFileName"] = convert_file

#         # 取得輸出的檔案目錄位置
#         convert_file_path = os.path.join(self.changed_folder_path, convert_file)

#         try:
#             if ShowRunningDataSwitch:
#                 print("convert_output")
#                 print(self.convert_output)

#             self.convert_output.to_csv(convert_file_path, index = False, encoding = "utf-8")

#             # 資訊寫入轉換結果 dict 中
#             self.change_result["Num"] = len(self.convert_output)

#             msg = f"檔案轉換完成，輸出檔名 {convert_file}。"
#             WChaLog("1", "ConvertSaleOutputInfo", self.dealer_id, self.file_name, msg)

#         except Exception as e:
#             msg = f"寫入檔案時發生錯誤。錯誤原因：{str(e)}"
#             WChaLog("2", "ConvertSaleOutputInfo", self.dealer_id, self.file_name, msg)
#             raise OSError(msg) from e

#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_output_info end.")

#     # 產出 sale 的 error report
#     def convert_error_output_info(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_error_output_info start.")
#         start_date, end_date = self.getMonthStartAndEndDate(self.file_name)
#         error_report_flag = False

#         if not self.not_in_master_file_data.empty:
#             error_report_flag = True
#             msg = "Masterfile 工作表中無此 Product ID。"
#             self.not_in_master_file_data.insert(0, "Exchange Error Issue", msg)
#             error_report_data = self.not_in_master_file_data

#         # 若 not_in_masterfile_data 為空，建立空白的error report DataFrame
#         else:
#             columns = ["Exchange Error Issue"] + self.input_file_data.columns.values
#             error_report_data = pd.DataFrame(columns = columns)

#         # sale convert error 有資料
#         if self.convert_error:
#             error_report_flag = True
#             row_index_list = list(self.convert_error.keys())
#             data = self.input_file_data.loc[row_index_list]
#             data.insert(0, "Exchange Error Issue", list(self.convert_error.values()))

#             # error report 資料合併
#             if error_report_data.empty:
#                 error_report_data = data

#             else:
#                 error_report_data = pd.concat([error_report_data, data], axis = 0)

#         error_report_data.insert(0, "Dealer Name", self.dealer_name)
#         error_report_data.insert(0, "Dealer ID", self.dealer_id)
#         # print(error_report_data)

#         if error_report_flag:
#             error_report_file_name = Config.SaleErrorReportFileName.replace\
#                 ("{DealerID}", str(self.dealer_id)).replace\
#                 ("{StartDate}", start_date).replace\
#                 ("{EndDate}", end_date)

#             error_report_folder_path = self.getErrorReportPath(self.dealer_id)
#             # 測試用
#             # error_report_folder_path = os.path.join("./datas")

#             try:
#                 if not os.path.exists(error_report_folder_path):
#                     os.makedirs(error_report_folder_path)
#                     msg = f"成功建立目錄 {error_report_folder_path} 。"
#                     WSysLog("1", "ConvertErrorOutputInfo", msg)

#                 if ShowRunningDataSwitch:
#                     print("error_report_data")
#                     print(error_report_data)

#                 # 將 transaction_date 欄位內容調整為：年/月/日
#                 error_report_data.loc[:, self.transaction_date] = self.changeTimeFormat\
#                     (error_report_data, self.transaction_date, "%Y/%m/%d")

#                 # 將 creation_date 欄位內容調整為：年/月/日
#                 error_report_data.loc[:, self.creation_date] = self.changeTimeFormat\
#                     (error_report_data, self.creation_date, "%Y/%m/%d")

#                 if ShowRunningDataSwitch:
#                     print("After format date data.")
#                     print("error_report_data")
#                     print(error_report_data)

#                 error_report_path = os.path.join(error_report_folder_path, error_report_file_name)
#                 error_report_data.to_excel(error_report_path, index = False)

#                 # 資訊寫入轉換結果 dict 中
#                 self.change_result["ErrorNum"] = len(error_report_data)
#                 self.change_result["ErrorReportFileName"] = error_report_file_name
#                 self.change_result["ErrorReportPath"] = error_report_path

#                 msg = f"Error檔案輸出完成，輸出檔名 {error_report_file_name}。"
#                 WChaLog("1", "ConvertErrorOutputInfo", self.dealer_id, self.file_name, msg)

#             except Exception as e:
#                 msg = f"輸出 Error Report 時發生錯誤。錯誤原因：{str(e)}"
#                 raise OSError(msg) from e

#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_error_output_info end.")

#     # 經銷商銷售檔案轉換成 BD 格式，sale轉換主程式
#     def changeSaleFile(self, dealer_id, file_name):
#         if ShowScheduleSwitch:
#             print("Function change_sale_file start.")

#         self.dealer_id = dealer_id
#         self.file_name = file_name

#         if ShowRunningDataSwitch:
#             print(f"\tdealer_id:{self.dealer_id}")
#             print(f"\tfile_name:{self.file_name}\n")

#         dealer_index = Config.DealerList.index(self.dealer_id)\
#             if self.dealer_id in Config.DealerList else None

#         if dealer_index is not None:
#             dealer_index += 1
#             if ShowRunningDataSwitch:
#                 print(f"dealer_index:{dealer_index}")

#             # 從 json 中取的經銷商資料
#             self.dealer_name = Config.DealerConfig[f"Dealer{dealer_index}"]["DealerName"]
#             self.dealer_country = Config.DealerConfig[f"Dealer{dealer_index}"]["DealerCountry"]
#             self.dealer_OUP_type = Config.DealerConfig\
#                 [f"Dealer{dealer_index}"]["SaleFile"]["OUPType"]

#             # 取得檔案的目錄位置
#             file_path = self.getFilePath(self.dealer_id, self.file_name)

#             # 測試用
#             # file_path = os.path.join("./datas", self.file_name)

#             # 讀取檔案內容，過濾附檔名、統一時間欄位內容格式、確認 pid 欄位之值無中文
#             self.input_file_data = self.checkProductIdValue(file_path)

#             if isinstance(self.input_file_data, bool):
#                 if ShowRunningDataSwitch:
#                     print("檔案附檔名不符合規範，系統將刪除該檔案。")
#                 self.delFile(file_path)

#             else:
#                 if ShowRunningDataSwitch:
#                     print("input_file_data")
#                     print(self.input_file_data)

#                 # 比對經銷商上傳的檔案，pid 值是否在 masterfile 存在
#                 self.in_master_file_data, self.not_in_master_file_data = \
#                     self.checkProdictIdInMasterFile(self.input_file_data, self.dealer_id)

#                 if ShowRunningDataSwitch:
#                     print("in_master_file_data")
#                     print(self.in_master_file_data)
#                     print()
#                     print("not_in_master_file_data")
#                     print(self.not_in_master_file_data)

#                 # 統計經銷商寫入的資料 pid 不在 masterfile 中的總數
#                 if len(self.not_in_master_file_data) > 0:
#                     self.change_status = "NO"
#                     msg = f"檔案中有 {len(self.not_in_master_file_data)} 筆資料在 master file 貨號中找不到。"
#                     WChaLog("2","ChangeSaleFile", self.dealer_id, self.file_name, msg)

#                 # 初始化 convert_output
#                 index_of_in_master_file_data = self.in_master_file_data["index"].to_list()
#                 self.convert_output = pd.DataFrame(index = index_of_in_master_file_data,
#                                                    columns = Config.SaleOutputFileHeader)

#                 # 轉換資料
#                 self.convert_data_in_rule()

#                 if ShowRunningDataSwitch:
#                     print("convert_output")
#                     print(self.convert_output)
#                     print()
#                     print("convert_error.keys")
#                     print(list(self.convert_error.keys()))
#                     print()
#                     print("convert_error")
#                     print(self.convert_error)

#                 # 將轉換過程中有錯誤的行從轉換輸出資料中移除
#                 if self.convert_error:
#                     self.convert_output.drop(index = list(self.convert_error.keys()), inplace=True)

#                 if ShowRunningDataSwitch:
#                     print("After drop error data.")
#                     print("convert_output")
#                     print(self.convert_output)
#                     print()
#                     print("convert_error")
#                     print(self.convert_error)
#                     print()

#                 # 轉換完成的結果寫入檔案
#                 self.convert_output_info()

#                 # 轉換錯誤的結果寫入Error Report
#                 self.convert_error_output_info()
#         else:
#             msg = f"在DealerList列表中搜尋不到此ID：{self.dealer_id}。"
#             WSysLog("2", "ChangeSaleFile", msg)

#         if ShowScheduleSwitch:
#             print("Function change_sale_file end.")

#         return self.change_result

# class InventoryDataChange(dataChangeFuncion):
#     """
#     庫存檔案轉換流程

#         class 內部函數呼叫使用
#             函數名稱                     說明
#             __init__                    class 內區域變數
#             convert_data_in_rule        根據庫存轉換規則轉換來源資料
#             get_file_date               以系統運作日期產出對應檔名
#             convert_output_info         產出 Inventory 轉換結果
#             convert_error_output_info   產出 Inventory 的 error report

#         外部呼叫函數
#             函數名稱                     說明
#             changeInventoryFile         庫存檔案轉換主流程；dealer_id -> str；file_name -> str
#     """

#     # class 內區域變數
#     def __init__(self):
#         super().__init__()
#         # 轉換規則
#         self.change_rule = Config.InventoryFileChangeRule

#         # 轉換檔案產出位置
#         self.changed_folder_path = Config.ChangeFolderPath
#         # 測試用
#         # self.changed_folder_path = "./datas"

#         # 設定預設值
#         self.dealer_id = None
#         self.dealer_name = None
#         self.file_name = None
#         self.input_file_data = None
#         self.in_master_file_data = None
#         self.not_in_master_file_data = None
#         self.convert_error = {}
#         self.change_status = "OK"
#         self.convert_output = None
#         self.change_result = {"Status" : self.change_status,
#                             "OutputFileName": None,
#                             "ErrorNum": 0,
#                             "Num" : 0,
#                             "ErrorReportFileName" : None,
#                             "ErrorReportPath" : None}

#     # 根據庫存轉換規則轉換來源資料
#     def convert_data_in_rule(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_data_in_rule start.")

#         input_data = self.in_master_file_data

#         for column_index in self.change_rule:
#             # 從 mapping.json 取出庫存檔案每個 col 對應的轉換參數
#             source_col = self.change_rule[column_index]["SourceName"]
#             target_col = self.change_rule[column_index]["ColumnName"]
#             rule = self.change_rule[column_index]["ChangeRule"]
#             value = self.change_rule[column_index]["Value"]

#             if ShowRunningDataSwitch:
#                 print(f"\t\trule:{rule}")

#             # 欄位為固定值
#             if rule == "FixedValue":
#                 if source_col == "DealerID":
#                     value = self.dealer_id
#                 self.convert_output[target_col] = self.fixedValue(value, len(input_data))

#             # 搬移資料
#             elif rule == "Move":
#                 self.convert_output[target_col] = self.moveRule(input_data, source_col)

#             # Quantity 特殊處理
#             elif rule == "MoveOrSearchUom":
#                 self.convert_output[target_col], no_value_in_masterfile = \
#                     self.moveOrSearchUom(input_data, source_col, self.dealer_id, target_col)
#                 self.convert_error.update(no_value_in_masterfile)

#             # Date Period 欄位為 Transaction Date 最後一天
#             elif rule == "LastTransactionDate":
#                 self.convert_output[target_col] = self.lastTransactionDate\
#                     (input_data, len(input_data))

#             else:
#                 self.change_status = "NO"
#                 msg = f"{rule} 此轉換規則不再範圍中。"
#                 WSysLog("3", "ConvertDataInRule", msg)
#                 raise ValueError(msg)

#             if ShowRunningDataSwitch:
#                 print(self.convert_output)

#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_data_in_rule end.")

#     # 以系統運作日期產出對應檔名
#     def get_file_date(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction get_file_date start.")
#             print("\t\tFunction get_file_date end.")
#         return str(datetime.now().year) + str(datetime.now().month) + str(datetime.now().day)

#     # 產出 Inventory 轉換結果
#     def convert_output_info(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_output_info start.")

#         self.convert_output.reset_index(drop = True)

#         # 取得系統運作日期
#         system_date = self.get_file_date()

#         # 各經銷商轉換出的庫存檔案檔名
#         convert_file_name = f"{self.dealer_id}_I_{system_date}.csv"

#         # 資訊寫入轉換結果 dict 中
#         self.change_result["OutputFileName"] = convert_file_name

#         # 取得輸出的檔案目錄位置
#         convert_file_path = os.path.join(self.changed_folder_path, convert_file_name)

#         try:
#             if ShowRunningDataSwitch:
#                 print("convert_output")
#                 print(self.convert_output)

#             self.convert_output.to_csv(convert_file_path, index=False)

#             # 資訊寫入轉換結果 dict 中
#             self.change_result["Num"] = len(self.convert_output)

#             msg = f"檔案轉換完成，輸出檔名 {convert_file_name}。"
#             WChaLog("1", "ConvertOutputInfo", self.dealer_id, self.file_name, msg)
#         except Exception as e:
#             msg = f"寫入檔案時發生錯誤。錯誤原因：{str(e)}"
#             WChaLog("2", "ConvertdOutputInfo", self.dealer_id, self.file_name, msg)
#             raise OSError(msg) from e

#         if ShowScheduleSwitch:
#             print("\t\tFunction Convert_output_info end.")

#     # 產出 Inventory 的 error report
#     def convert_error_output_info(self):
#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_error_output_info start.")

#         error_report_flag = False

#         # 取得系統運作日期
#         start_date, end_date = self.getMonthStartAndEndDate(self.file_name)

#         if not self.not_in_master_file_data.empty:
#             error_report_flag = True
#             msg = "Masterfile 工作表中無此 Product ID。"
#             self.not_in_master_file_data.insert(0, "Exchange Error Issue", msg)
#             error_report_data = self.not_in_master_file_data

#         # 若 not_in_masterfile_data 為空，建立空白的error report DataFrame
#         else:
#             columns = ["Exchange Error Issue"] + self.input_file_data.columns.values
#             error_report_data = pd.DataFrame(columns=columns)

#         # inventory convert error 有資料
#         if self.convert_error:
#             error_report_flag = True
#             row_index_list = list(self.convert_error.keys())
#             data = self.input_file_data.loc[row_index_list]
#             data.insert(0, "Exchange Error Issue", list(self.convert_error.values()))

#             # error report 資料合併
#             if error_report_data.empty:
#                 error_report_data = data

#             else:
#                 error_report_data = pd.concat([error_report_data, data], axis = 0)

#         error_report_data.insert(0, "Dealer Name", self.dealer_name)
#         error_report_data.insert(0, "Dealer ID", self.dealer_id)
#         # print(error_report_data)

#         if error_report_flag:
#             error_report_file_name = Config.InventoryErrorReportFileName.replace\
#                 ("{DealerID}", str(self.dealer_id)).replace\
#                 ("{StartDate}", start_date).replace\
#                 ("{EndDate}", end_date)

#             error_report_folder_path = self.getErrorReportPath(self.dealer_id)
#             # 測試用
#             # error_report_folder_path = os.path.join("./datas")

#             try:
#                 if not os.path.exists(error_report_folder_path):
#                     os.makedirs(error_report_folder_path)
#                     msg = f"成功建立目錄 {error_report_folder_path} 。"
#                     WSysLog("1", "ConvertdErrorOutputInfo", msg)

#                 if ShowRunningDataSwitch:
#                     print("error_report_data")
#                     print(error_report_data)

#                 # 將 transaction_date 欄位內容調整為：年/月/日
#                 error_report_data.loc[:, self.transaction_date] = self.changeTimeFormat\
#                     (error_report_data, self.transaction_date, "%Y/%m/%d")

#                 # 將 creation_date 欄位內容調整為：年/月/日
#                 error_report_data.loc[:, self.creation_date] = self.changeTimeFormat\
#                     (error_report_data, self.creation_date, "%Y/%m/%d")

#                 if ShowRunningDataSwitch:
#                     print("After format date data.")
#                     print("error_report_data")
#                     print(error_report_data)

#                 error_report_path = os.path.join(error_report_folder_path, error_report_file_name)
#                 error_report_data.to_excel(error_report_path, index = False)

#                 # 資訊寫入轉換結果 dict 中
#                 self.change_result["ErrorNum"] = len(error_report_data)
#                 self.change_result["ErrorReportFileName"] = error_report_file_name
#                 self.change_result["ErrorReportPath"] = error_report_path

#                 msg = f"Error檔案輸出完成，輸出檔名 {error_report_file_name}。"
#                 WChaLog("1", "ConvertdErrorOutputInfo", self.dealer_id, self.file_name, msg)

#             except Exception as e:
#                 msg = f"輸出 Error Report 時發生錯誤。錯誤原因：{str(e)}"
#                 raise OSError(msg) from e

#         if ShowScheduleSwitch:
#             print("\t\tFunction convert_error_output_info end.")

#     # 經銷商庫存檔案轉換成 BD 格式，Inventory轉換主程式
#     def changeInventoryFile(self, dealer_id, file_name):
#         if ShowScheduleSwitch:
#             print("Function change_inventory_file start.")

#         self.dealer_id = dealer_id
#         self.file_name = file_name

#         if ShowRunningDataSwitch:
#             print(f"\tdealer_id:{self.dealer_id}")
#             print(f"\tfile_name:{self.file_name}\n")

#         dealer_index = Config.DealerList.index(self.dealer_id)\
#             if self.dealer_id in Config.DealerList else None

#         if dealer_index is not None:
#             dealer_index += 1
#             if ShowRunningDataSwitch:
#                 print(f"dealer_index:{dealer_index}")

#             # 從 json 中取的經銷商資料
#             self.dealer_name = Config.DealerConfig[f"Dealer{dealer_index}"]["DealerName"]

#             # 取得檔案的目錄位置
#             file_path = self.getFilePath(self.dealer_id, self.file_name)

#             # 測試用
#             # file_path = os.path.join("./datas", self.file_name)

#             # 讀取檔案內容，過濾附檔名、統一時間欄位內容格式、確認 pid 欄位之值無中文
#             self.input_file_data = self.checkProductIdValue(file_path)

#             if isinstance(self.input_file_data, bool):
#                 if ShowRunningDataSwitch:
#                     print("檔案附檔名不符合規範，系統將刪除該檔案。")
#                 self.delFile(file_path)

#             else:
#                 if ShowRunningDataSwitch:
#                     print("input_file_data")
#                     print(self.input_file_data)

#                 # 比對經銷商上傳的檔案，pid 值是否在 masterfile 存在
#                 self.in_master_file_data, self.not_in_master_file_data = \
#                     self.checkProdictIdInMasterFile(self.input_file_data, self.dealer_id)

#                 if ShowRunningDataSwitch:
#                     print("in_master_file_data")
#                     print(self.in_master_file_data)
#                     print()
#                     print("not_in_master_file_data")
#                     print(self.not_in_master_file_data)

#                 # 統計經銷商寫入的資料 pid 不在 masterfile 中的總數
#                 if len(self.not_in_master_file_data) > 0:
#                     self.change_status = "NO"
#                     msg = f"檔案中有 {len(self.not_in_master_file_data)} 筆資料在 master file 貨號中找不到。"
#                     WChaLog("2","ChangeInventoryFile", self.dealer_id, self.file_name, msg)

#                 # 初始化 convert_output
#                 index_of_in_master_file_data = self.in_master_file_data["index"].to_list()
#                 self.convert_output = pd.DataFrame(index = index_of_in_master_file_data,
#                                                    columns = Config.InventoryOutputFileHeader)

#                 # 轉換資料
#                 self.convert_data_in_rule()

#                 if ShowRunningDataSwitch:
#                     print("convert_output")
#                     print(self.convert_output)
#                     print()
#                     print("convert_error.keys")
#                     print(list(self.convert_error.keys()))
#                     print()
#                     print("convert_error")
#                     print(self.convert_error)

#                 # 將轉換過程中有錯誤的行從轉換輸出資料中移除
#                 if self.convert_error:
#                     self.convert_output.drop(index = list(self.convert_error.keys()), inplace=True)

#                 if ShowRunningDataSwitch:
#                     print("After drop error data.")
#                     print("convert_output")
#                     print(self.convert_output)
#                     print()
#                     print("convert_error")
#                     print(self.convert_error)
#                     print()

#                 # 轉換完成的結果寫入檔案
#                 self.convert_output_info()

#                 # 轉換錯誤的結果寫入Error Report
#                 self.convert_error_output_info()

#         if ShowScheduleSwitch:
#             print("Function change_inventory_file end.")

#         return self.change_result

# # 轉換經銷商上傳檔案流程
# def Changing(check_right_list):
#     for dealer_id in Config.DealerList:
#         error_files, error_paths = [], []
#         dealer_path = os.path.join(Config.DealerFolderPath, dealer_id)

#         file_names = [file for file in os.listdir(dealer_path)\
#                       if os.path.isfile(os.path.join(dealer_path, file))]

#         for file_name in file_names:
#             file_type, _ = get_file_type(dealer_id, file_name)

#             if file_type == "Sale":
#                 # class初始化
#                 sale_change = SaleDataChange()
#                 change_result = sale_change.changeSaleFile(dealer_id, file_name)
#                 # print("ChangeResult")
#                 # print(change_result)

#                 status = change_result["Status"]
#                 output_file_name = change_result["OutputFileName"]
#                 error_num = change_result["ErrorNum"]
#                 num = change_result["Num"]
#                 error_report = change_result["ErrorReportFileName"]
#                 error_report_path = change_result["ErrorReportPath"]

#                 if error_num != 0:
#                     error_files.append(error_report)
#                     error_paths.append(error_report_path)

#             else:
#                 # class初始化
#                 inventory_change = InventoryDataChange()
#                 change_result = inventory_change.changeInventoryFile(dealer_id, file_name)
#                 # print("ChangeResult")
#                 # print(change_result)

#                 status = change_result["Status"]
#                 output_file_name = change_result["OutputFileName"]
#                 error_num = change_result["ErrorNum"]
#                 num = change_result["Num"]
#                 error_report = change_result["ErrorReportFileName"]
#                 error_report_path = change_result["ErrorReportPath"]

#                 if error_num != 0:
#                     error_files.append(error_report)
#                     error_paths.append(error_report_path)

#             for data_id, name in check_right_list.items():
#                 if file_name == name:
#                     write_data = {
#                         "ChangeData":{
#                             "ID":data_id,
#                             "轉換狀態":status,
#                             "轉換後檔案名稱":output_file_name,
#                             "轉換錯誤筆數":error_num,
#                             "轉換後總筆數":num
#                         }
#                     }
#                     WriteSubRawData(write_data)
#                     print("WriteData")
#                     print(write_data)
#                     break

#         if error_files:
#             mail_data = {"ErrorReportFileName" : "、".join(error_files)}
#             send_info = {"Mode" : "ErrorReport",
#                         "DealerID" : dealer_id,
#                         "MailData" : mail_data,
#                         "FilesPath" : error_paths}
#             SendMail(send_info)
#             print("SendInfo")
#             print(send_info)

# # 合併 Inventory 檔案
# def MergeInventoryFile():
#     file_names = [file for file in os.listdir(Config.ChangeFolderPath) \
#         if os.path.isfile(os.path.join(Config.ChangeFolderPath, file))]

#     # 取得要合併的檔案
#     file_list, data_dates = [], []
#     for file_name in file_names:
#         part = re.split(r"[._]" ,file_name)
#         if part[1] == "I":
#             file_list.append(file_name)
#             data_dates.append(part[2])

#     if file_list:
#         # 取出檔名中的檔案日期
#         data_date = data_dates[0]

#         # 合併檔案
#         dataframes = [pd.read_csv(os.path.join(Config.ChangeFolderPath, file))\
#             for file in file_list]
#         combined_df = pd.concat(dataframes, ignore_index=True)

#         # 輸出檔案
#         changed_file_name = Config.InventoryOutputFileName.replace\
#             ("{CountryCode}", Config.InventoryOutputFileCountryCode).replace\
#             ("{LastDate}", data_date)

#         try:
#             # 輸出當日合併的庫存總數據
#             combined_df.to_csv(os.path.join(Config.ChangeFolderPath,
#                 f"{changed_file_name}.{Config.InventoryOutputFileExtension}"),
#                 sep = ",", index=False)

#             msg = f"{len(file_list)} 份檔案成功合併成 {changed_file_name}.{Config.InventoryOutputFileExtension}。"
#             WSysLog("1", "MargeInventoryFile", msg)

#             # 搬移經銷商庫存轉換後的檔案
#             for dealer_id in Config.DealerList:
#                 target_folder = os.path.join(Config.ChangeFolderPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))
#                 if not os.path.exists(target_folder):
#                     os.makedirs(target_folder)

#                 for file in file_list:
#                     part = re.split(r"[._]" ,file)
#                     if (part[0] == dealer_id) and (part[1] == "I"):
#                         file_source = os.path.join(Config.ChangeFolderPath, file)
#                         file_target = os.path.join(target_folder, file)
#                         shutil.move(file_source, file_target)

#                         if os.path.exists(file_target):
#                             msg = f"檔案搬移至 {target_folder} 成功"
#                             WSysLog("1", "MoveInventoryFile", msg)
#                         else:
#                             msg = f"檔案搬移至 {target_folder} 失敗"
#                             WSysLog("2", "MoveInventoryFile", msg)

#         except Exception as e:
#             msg = f"合併檔案發生未知錯誤，{e}。"
#             WSysLog("3", "MargeInventoryFile", msg)

# # 檔案上傳EFT雲端完成後歸檔
# def FileArchiving():
#     file_names = [file for file in os.listdir(Config.ChangeFolderPath) \
#         if os.path.isfile(os.path.join(Config.ChangeFolderPath, file))]

#     for file in file_names:
#         part = file.split("_")
#         for dealer_id in Config.DealerList:
#             if part[0] == dealer_id:
#                 target_folder = os.path.join(Config.ChangeFolderPath, dealer_id, datetime.strftime(Config.SystemTime, "%Y%m"))

#                 # 確認目標目錄是否存在
#                 if not os.path.exists(target_folder):
#                     os.makedirs(target_folder)

#                 file_source = os.path.join(Config.ChangeFolderPath, file)
#                 file_target = os.path.join(target_folder, file)
#                 shutil.move(file_source, file_target)

#                 if os.path.exists(file_target):
#                     msg = f"檔案搬移至 {target_folder} 成功"
#                     WSysLog("1", "MoveInventoryFile", msg)

#                 else:
#                     msg = f"檔案搬移至 {target_folder} 失敗"
#                     WSysLog("2", "MoveInventoryFile", msg)

#         # 合併後的庫存總表
#         if part[0] == Config.InventoryOutputFileCountryCode:
#             target_folder = os.path.join(Config.ChangeFolderPath, Config.MergeInventoryFolder, datetime.strftime(Config.SystemTime, "%Y%m"))

#             # 確認目標目錄是否存在
#             if not os.path.exists(target_folder):
#                 os.makedirs(target_folder)

#             file_source = os.path.join(Config.ChangeFolderPath, file)
#             file_target = os.path.join(target_folder, file)
#             shutil.move(file_source, file_target)

#             if os.path.exists(file_target):
#                 msg = f"檔案搬移至 {target_folder} 成功"
#                 WSysLog("1", "MoveInventoryFile", msg)

#             else:
#                 msg = f"檔案搬移至 {target_folder} 失敗"
#                 WSysLog("2", "MoveInventoryFile", msg)
'''