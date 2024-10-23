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

"""
改寫Mapping檔案
"""
import os
import math
import pandas as pd
from dateutil import parser

from Log import WSysLog, WChaLog
from Config import AppConfig

Config = AppConfig()

class dataChangeFuncion:
    """
    MasterFile Header：
    經銷商號碼\nSold-to code, 貨號,	UOM, std cost (EA),	DP(EA),	KADP(EA), IVY (EA),	起,	迄

    KAList Header：
    經銷商ID, 客戶號, 起, 迄, Price TYPE

    檔案轉換方式方法
    """
    # class內區域變數
    def __init__(self):
        # class內通用參數
        self.master_file_header = ["經銷商號碼\nSold-to code",
                                   "貨號",
                                   "UOM",
                                   "std cost (EA)",
                                   "DP(EA)","KADP(EA)",
                                   "IVY (EA)",
                                   "起",
                                   "迄"]

        self.kalist_file_header = [ "經銷商ID",
                                    "客戶號",
                                    "起",
                                    "迄",
                                    "Price TYPE"]

        self.product_id = "Product ID"
        self.transaction_date = "Transaction Date"
        self.creation_date = "Creation Date"
        self.buyer_id = "Buyer ID"

        date_cols = ["起", "迄"]
        master_folder_path = "./datas"
        master_file_name = "MasterFile.xlsx"
        master_file_sheet_name = "MasterFile"
        kalist_file_sheet_name = "KAList"

        master_file_path = os.path.join(master_folder_path, master_file_name)

        try:
            # MasterFile 資料
            self.master_file_data = pd.read_excel(master_file_path,
                                                sheet_name = master_file_sheet_name,
                                                dtype = str)

            # 比對 MasterFile 工作表要與規定的符合
            master_file_data_header = self.master_file_data.columns.values
            if set(self.master_file_header) != set(master_file_data_header):
                msg = "MasterFile工作表的表頭與規定值不匹配。"
                raise ValueError(msg)

            for col in date_cols:
                self.master_file_data[col] = pd.to_datetime(self.master_file_data[col],
                                                            format = "%Y%m%d")

            # KAList 資料
            self.kalist_file_data = pd.read_excel(master_file_path,
                                                sheet_name = kalist_file_sheet_name,
                                                dtype = str)

            # 比對 KAList 工作表要與規定的符合
            kalist_file_data_header = self.kalist_file_data.columns.values
            if set(self.kalist_file_header) != set(kalist_file_data_header):
                msg = "KAList工作表的表頭與規定值不匹配。"
                raise ValueError(msg)

            for col in date_cols:
                self.kalist_file_data[col] = pd.to_datetime(self.kalist_file_data[col],
                                                            format = "%Y%m%d")

            msg = "成功讀取 MasterFile 資料。"
            WSysLog("1", "ReadMasterFile", msg)

        except Exception as e:
            msg = f"讀取 MasterFile時發生錯誤。錯誤原因：{str(e)}"
            WSysLog("3", "ReadMasterFile", msg)
            raise FileNotFoundError(msg) from e

    # 將 master_file 內容輸出
    def print_master_file_data(self):
        print("Function master_file_test start.")
        print(self.master_file_data)
        print("Function master_file_test finish.")

    # 將 kalist 內容輸出
    def print_kalist_file_data(self):
        print("Function kalist_file_test start.")
        print(self.kalist_file_data)
        print("Function kalist_file_test finish.")

    # 統一時間欄位內容格式
    def parse_and_format_date(self, date_str, output_format = "%Y/%m/%d"):
        # print("Function parse_and_format_date start.")
        try:
            parsed_date = parser.parse(date_str)
            return parsed_date.strftime(output_format)
        except (ValueError, TypeError):
            return date_str
        # finally:
        #     print("Function parse_and_format_date finish.")

    # 確認檔案 product_id 欄位值非中文
    def checkProductIdValue(self, file_path):
        print("Function check_product_id_value start.")
        file_name = os.path.basename(file_path)
        _, file_extension = os.path.splitext(file_name)
        file_extension = file_extension.lower()

        if file_extension in Config.AllowFileExtensions:
            file_data = pd.read_csv(file_path, dtype = str)\
                if file_extension == ".csv"\
                else pd.read_excel(file_path, dtype = str)

            # print(file_data)

            # 將時間欄位內容資料型態統一
            date_cols = [self.transaction_date, self.creation_date]
            for col in date_cols:
                file_data[col] = file_data[col].apply(lambda x: self.parse_and_format_date(str(x)))
                file_data[col] = pd.to_datetime(file_data[col],
                                                format = "%Y/%m/%d",
                                                errors = "coerce")
            file_data[self.product_id] = file_data[self.product_id].astype(str).str.strip()

            # 過濾 Product ID 僅允許 a-z, A-Z, 0-9, - 符號
            file_data = file_data[file_data[self.product_id].str.contains\
                ("^[a-zA-Z0-9-]+$", regex=True, na=False)]

            # 刷新index值
            file_data = file_data.reset_index()
            print("Function check_product_id_value end.")
            return file_data

        else:
            print("Function check_product_id_value end.")
            return False

    # 比對檔案中的 product id 存在於 master file 中
    def checkProdictIdInMasterFile(self, input_file_data, dealer_id):
        print("Function check_prodict_id_in_master_file start.")
        col_dealer_id = self.master_file_header[0]
        col_product_id = self.master_file_header[1]
        pid_not_in_master_file = []
        # input_file_data -> dataFrame
        # 取出輸入資料中的 "product id" 欄位資料，去掉重複值，排序
        input_data_product_id = input_file_data.loc[:, self.product_id].tolist()
        input_data_product_id = sorted(list(set(input_data_product_id)))
        # print(input_data_product_id)
        # print(len(input_data_product_id))

        # master_file_data -> dataFrame，此處輸入的masterfile是該經銷商的資料，非整份masterfile
        # 取出輸入資料中的 "貨號" 欄位資料，去掉重複值，排序
        master_file_data = self.master_file_data[\
            self.master_file_data[col_dealer_id] == dealer_id]
        master_file_product_id = master_file_data.loc[:, col_product_id].tolist()
        master_file_product_id = sorted(list(set(master_file_product_id)))

        # 使用迴圈檢查經銷商上傳檔案中的 product_id 是否存在於 master_file 資料中
        for pid in input_data_product_id:
            if pid not in master_file_product_id:
                pid_not_in_master_file.append(pid)
        # print(f"pid_not_in_master_file:{pid_not_in_master_file}")

        input_data_in_master_file = input_file_data[\
            ~input_file_data[self.product_id].isin(pid_not_in_master_file)]
        # print("input_data_in_master_file")
        # print(input_data_in_master_file)

        input_data_not_in_master_file = input_file_data[\
            input_file_data[self.product_id].isin(pid_not_in_master_file)]
        # print("input_data_not_in_master_file")
        # print(input_data_not_in_master_file)
        print("Function check_prodict_id_in_master_file end.")
        return input_data_in_master_file, input_data_not_in_master_file

    # 將原先欄位的值移動到搬移到新的欄位
    def moveRule(self, input_data, input_col):
        print("Function move_rule start.")
        print("Function move_rule end.")
        return input_data[input_col]

    # 欄位值固定為某些值
    def fixedValue(self, value, row):
        print("Function fixed_value start.")
        print("Function fixed_value end.")
        return [value] * row

    # 轉換欄位內容的時間格式
    def changeTimeFormat(self, input_data, input_col, date_format):
        print("Function change_time_format start.")
        print("Function change_time_format end.")
        return input_data[input_col].dt.strftime(date_format)

    # 在MasterFile中搜尋產品ID，回傳需要的 value 值
    def search_pid_in_master_file(self, dealer_id, product_id, data_date):
        print("Function search_pid_in_master_file start.")
        col_dealer_id = self.master_file_header[0]
        col_product_id = self.master_file_header[1]
        col_start_date = self.master_file_header[7]
        col_end_date = self.master_file_header[8]

        # 從 master_file 中取出對應 經銷商ID 的資料
        master_file_data = self.master_file_data\
            [self.master_file_data[col_dealer_id] == dealer_id]

        search_pid_data = master_file_data[master_file_data[col_product_id] == product_id]

        if not search_pid_data.empty:
            # 透過data_date篩選除對應區間的masterfile資料
            pid_data_in_date = search_pid_data[(search_pid_data[col_start_date] <= data_date) &
                                                (search_pid_data[col_end_date] >= data_date)]
            if not pid_data_in_date.empty:
                print("Function search_pid_in_master_file end.")
                return True, pid_data_in_date

            else:
                msg = f"經銷商：{dealer_id} 的產品ID：{product_id}，在 masterfile 工作表中搜尋不到對應的起迄區間。"
                WSysLog("3", "SearchPidInMasterFile", msg)
                msg = "在 masterfile 檔案中搜尋不到對應的起迄區間。"
                print("Function search_pid_in_master_file end.")
                return None, msg
        else:
            msg = f"經銷商：{dealer_id} 的產品ID：{product_id}，在 masterfile 工作表中搜尋不到。"
            WSysLog("3", "SearchPidInMasterFile", msg)
            msg = "在 masterfile 檔案中搜尋不到此 Product ID。"
            print("Function search_pid_in_master_file end.")
            return False, msg

    # 搬移或是轉換 Uom 值
    def moveOrSearchUom(self, input_data, source_col, dealer_id, target_col):
        print("Function move_or_search_uom start.")
        col_uom = self.master_file_header[2]
        not_get_value_row, not_get_value_msg = [], []

        value_list_in_source = input_data[source_col].tolist()

        # 篩選來源col欄位中值為空白的row，並取得index
        source_na_index_list = input_data[input_data[source_col].isna()].index.tolist()

        for row in source_na_index_list:
            # print(row)
            input_data_product_id = input_data.loc[row, self.product_id]
            input_data_date = input_data.loc[row, self.transaction_date]
            # print(f"input_data_product_id:{input_data_product_id}")
            # print(f"input_data_date:{input_data_date}")
            search_result, pid_data_in_date = self.search_pid_in_master_file\
                (dealer_id, input_data_product_id, input_data_date)

            if search_result:
                target_col_list = pid_data_in_date[col_uom].to_list()
                try:
                    value = target_col_list[-1]
                    if (isinstance(value, float)) and (math.isnan(value)):
                        row_in_masterfile = pid_data_in_date.iloc[-1].name + 2
                        msg = f"masterfile檔案中 {col_uom} 欄位第 {row_in_masterfile} 行數值為空。"
                        WSysLog("2", "SearchPidInMasterFile", msg)
                        not_get_value_row.append(row)

                    elif isinstance(value, str):
                        uom_in_masterfile = int(value)
                        # print(input_data.loc[row, target_col])
                        # print(type(input_data.loc[row, target_col]))
                        if (isinstance(input_data.loc[row, target_col], float)) and\
                            (math.isnan(input_data.loc[row, target_col])):
                            msg = f"經銷商寫入的 {target_col} 欄位於第 {row + 2} 數值為空。"
                            WSysLog("2", "SearchPidInMasterFile", msg)
                            not_get_value_row.append(row)

                        else:
                            changed_value = uom_in_masterfile *\
                                float(input_data.loc[row, target_col])
                            value_list_in_source[row] = str(changed_value)

                except Exception as e:
                    msg = f"將搜尋的資訊轉換為數值時發生錯誤。錯誤原因：{str(e)}"
                    WSysLog("3", "SearchPidInMasterFile", msg)
                    raise TypeError(msg) from e

            else:
                not_get_value_row.append(row)
                not_get_value_msg.append(pid_data_in_date)

        # print("not_get_value_row")
        # print(not_get_value_row)
        for index in sorted(not_get_value_row, reverse=True):
            del value_list_in_source[index]

        print("Function move_or_search_uom end.")
        return value_list_in_source, not_get_value_row, not_get_value_msg

    def get_dp_type_in_kalist(self, dealer_id, buyer_id, data_date):
        print("Function get_dp_type start.")
        col_dealer_id_in_ka = self.kalist_file_header[0]
        col_buyer_id_in_ka = self.kalist_file_header[1]
        col_start_date = self.kalist_file_header[2]
        col_end_date = self.kalist_file_header[3]
        col_price_type = self.kalist_file_header[4]
        price_type = "DP"

        # 從 kalist 工作表中篩選經銷商與客戶號資訊
        search_data_in_ka = self.kalist_file_data\
            [(self.kalist_file_data[col_dealer_id_in_ka] == dealer_id) &
            (self.kalist_file_data[col_buyer_id_in_ka] == buyer_id)]

        if not search_data_in_ka.empty:
            buyer_data_in_date = search_data_in_ka\
                [(search_data_in_ka[col_start_date] <= data_date) &
                (search_data_in_ka[col_end_date] >= data_date)]

            if not buyer_data_in_date.empty:
                type_list = buyer_data_in_date[col_price_type].to_list()
                price_type = type_list[-1]
                print("Function get_dp_type end.")
                return price_type
            else:
                msg = f"經銷商ID： {dealer_id} 的客戶號： {buyer_id} 資料，在 KAList 工作表中未搜尋到符合時間區間的資料。"
                WSysLog("2", "get_dp_type_in_kalist", msg)
                print("Function get_dp_type end.")
                return price_type
        else:
            msg = f"經銷商ID： {dealer_id} 的客戶號： {buyer_id} 資料，在 KAList 工作表中搜尋不到 。"
            WSysLog("2", "get_dp_type_in_kalist", msg)
            print("Function get_dp_type end.")
            return price_type

    # 搜尋 Dp 資料，若在 kalist 中，則需篩選對應的 value
    def searchDP(self, input_data, dealer_id):
        print("Function search_dp start.")
        # col_dealer_id_in_ka = self.kalist_file_header[0]
        # col_buyer_id_in_ka = self.kalist_file_header[1]
        if dealer_id in Config.KADealerList:
            for row in range(len(input_data)):
                # print(row)
                # print(type(row))
                buyer_id = input_data.loc[row, self.buyer_id]
                data_date = input_data.loc[row, self.transaction_date]
                self.get_dp_type_in_kalist(dealer_id, buyer_id, data_date)
                if row == 5:
                    break
        # for row in range(len(input_data)):
        print("Function search_dp end.")

# class dealerDataChange(dataChangeFuncion):
#     def changeSaleFile():
#         print()
#     def changeInventoryFile():
#         print()

def test(data):
    change_function = dataChangeFuncion()
    # change_function.print_kalist_file_data()
    data = change_function.checkProductIdValue(data)
    # print(data)
    # change_function.check_prodict_id_in_master_file(data, "1002317244")
    # change_function.moveOrSearchUom(data, "Original Quantity", "1002317244", "Quantity")
    # prodict_id = data["Product ID"]
    # print(prodict_id)
    change_function.searchDP(data, "1002317244")

if __name__ == "__main__":
    test_data_file_name = "Unimed_S_202410152008_YTD.xls"
    test_data_path = "./datas"
    test_data_file_path = os.path.join(test_data_path, test_data_file_name)
    # print(test_data_file_path)
    # data = pd.read_excel(test_data_file_path, dtype = str)
    # print(data)
    test(test_data_file_path)
