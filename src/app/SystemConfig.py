# -*- coding: utf-8 -*-

'''
檔案說明：讀取、變更 ./src/config 底下的 json 檔案
Writer：Qian
'''

import os, json, pandas as pd
from __init__ import ConfigJsonFile

ConfigInfo = ConfigJsonFile()

# 讀取Json檔案
def read_json(file_path):
    with open(file_path, "r", encoding = "UTF-8") as file:
        data = json.load(file)
    return data

# 讀取 system.json
def Config():
    return read_json(ConfigInfo.ConfigPath)

# 讀取 dealer.json
def DealerConf():
    return read_json(ConfigInfo.DealerPath)

# 讀取 check_rule.json
def CheckRule():
    return read_json(ConfigInfo.CheckPath)

# 讀取 dealer_format.json
def DealerFormatConf():
    return read_json(ConfigInfo.DealerFormatPath)

# 讀取 mapping_rule.json
def MappingRule():
    return read_json(ConfigInfo.MappingPath)

# 讀取 mail.json
def MailRule():
    return read_json(ConfigInfo.MailRulePath)

# 讀取 user.json
def User():
    return read_json(ConfigInfo.UserConfigPath)

# 讀取 files.json
def File():
    return read_json(ConfigInfo.FileConfigPath)

# 將 excel 內容轉變為json ##
def write_rule_json(file_path, sheet, data_name, output_name):
    df = pd.read_excel(file_path,sheet_name=sheet)
    data_list = df.to_dict("records")
    final_json = {
        data_name: data_list
    }
    output_path = os.path.join("./src/config", output_name)
    with open(output_path, "w",encoding="UTF-8") as f:
        json.dump(final_json, f, ensure_ascii=False, indent=2)

# 合併json檔案 ##
def marge_json_files(file_path, file1_name, file2_name, output_name):
    File1Path = os.path.join(file_path, file1_name)
    File2Path = os.path.join(file_path, file2_name)
    OutputPath = os.path.join(file_path, output_name)
    with open(File1Path, "r", encoding = "UTF-8") as f1:
        data1 = json.load(f1)
    with open(File2Path, "r", encoding = "UTF-8") as f2:
        data2 = json.load(f2)
    merged_data = {**data1, **data2}
    with open(OutputPath, "w", encoding = "UTF-8") as of:
        json.dump(merged_data, of, ensure_ascii = False, indent = 2)
    try:
        os.remove(File1Path)
        os.remove(File2Path)
        print("finish")
    except OSError as e:
        print(f"error: {e}")

# 將 maping rule 的excel轉變為json ##
def MakeMappingRuleJson():
    DataName = ["Sale","Inventory"]
    FilePath = "./docs/data_format/DataFormatFromBD.xlsx"
    Sheet = "FileMappingRule"
    Output = "MappingRuleOutput.json"
    for i in DataName:
        SheetName = i + Sheet
        OutputName = i + Output
        write_rule_json(FilePath, SheetName, i, OutputName)
    File1Name = str(DataName[0]) + Output
    File2Name = str(DataName[1]) + Output
    marge_json_files(ConfigInfo.ConfigPath, File1Name, File2Name, ConfigInfo.MappingPath)

# 將經銷商資訊轉成 json ##
def DealerJson():
    output_path = os.path.join(ConfigInfo.ConfigPath, ConfigInfo.DealerPath)
    summary_sheet_name = "Dealer Summary"
    file_path = "./docs/dealer/BD合作之經銷商資料.xlsx"
    contact_list = ["Contact1","Contact2","ContactProject"]
    format_list = ["PaymentCycle","KeyWord","Extension","FileHeader"]
    part1_list = ["DealerID","DealerCompiled","DealerName","DealerKind","TelephoneNumber"]
    part2_list = ["Position","Name","Mail","Ex"]
    part3_list = ["Sale File Payment Cycle","Sale File Key","Sale File Extension"]
    part4_list = ["Inventory File Payment Cycle","Inventory File Key","Inventory File Extension"]

    df = pd.read_excel(file_path, sheet_name = summary_sheet_name)
    dealer_id = df["DealerID"].dropna().reset_index(drop = True).astype(int)
    dealer_id = dealer_id.apply(str).to_list()
    OutputData = {"DealerList":dealer_id}

    for i in range(len(dealer_id)):
        j = i*3
        part1 = df.loc[j:(j+2),part1_list]
        part2 = df.loc[j:(j+2),part2_list]
        part3 = df.loc[j:(j+2),part3_list]
        part4 = df.loc[j:(j+2),part4_list]

        part1 = part1.iloc[0]
        part3 = part3.iloc[[0]]
        part4 = part4.iloc[[0]]
        part1["DealerID"] = str(int(part1["DealerID"]))
        part1_dic = part1.to_dict()

        Contact ={}
        for k in range(len(part2)):
            part2_exchang = part2.iloc[[k]]
            new_columns = [contact_list[k] + col for col in part2_list]
            part2_exchang.columns = new_columns
            part2_exc = part2_exchang.iloc[0]
            part2_dic = part2_exc.to_dict()
            Contact.update(part2_dic)

        SaleSheetName = f"Dealer{i+1}_Sale"
        InventorySheetName = f"Dealer{i+1}_Inventory"
        SaleData = pd.read_excel(file_path, sheet_name = SaleSheetName)
        InventoryData = pd.read_excel(file_path, sheet_name = InventorySheetName)
        sale_header = SaleData.columns.to_list()
        inventory_header = InventoryData.columns.to_list()

        part3.columns = format_list[:3]
        part4.columns = format_list[:3]
        part3_exchang = part3.iloc[0]
        part4_exchang = part4.iloc[0]
        part3_dic = part3_exchang.to_dict()
        part4_dic = part4_exchang.to_dict()
        part3_dic[format_list[3]] = sale_header
        part4_dic[format_list[3]] = inventory_header

        data = {}
        data.update(part1_dic)
        data.update(Contact)
        data["SaleFile"] = part3_dic
        data["InventoryFile"] = part4_dic
        OutputData[f"Dealer{i+1}"] = data

    with open(output_path, "w",encoding="UTF-8") as f:
        json.dump(OutputData, f, ensure_ascii=False, indent=2)

# 取得 excel column 名稱
def excel_column_name(n):
    result = []
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result.append(chr(65 + remainder))
    return "".join(result[::-1])

# 抓取經銷商 Header，跟默認的比對 ##
def HeaderChange():
    write_data = {}
    dealer_format_config = DealerFormatConf()
    dealer_config = DealerConf()
    DealerList = dealer_config["DealerList"]
    sale_format = dealer_format_config["Defualt"]["SaleFileHeader"]
    inventory_format = dealer_format_config["Defualt"]["InventoryFileHeader"]
    for dealer_id in range(len(DealerList)):
        sale_flag, inve_flag = True, True
        index = dealer_id + 1
        sale_header = dealer_config[f"Dealer{index}"]["SaleFile"]["FileHeader"]
        inventory_header = dealer_config[f"Dealer{index}"]["InventoryFile"]["FileHeader"]
        location_sale, location_inve , location_sale_index, location_inve_index = [], [], [], []

        for i in range(len(sale_format)):
            hf = sale_format[i]
            hf_lower = hf.replace(" ", "").lower()
            for j in range(len(sale_header)):
                h = sale_header[j]
                h_lower = h.replace(" ", "").lower()
                if hf_lower == h_lower:
                    location_sale.append(excel_column_name(j+1))
                    location_sale_index.append(j)
                    break

        for i in range(len(inventory_format)):
            hf = inventory_format[i]
            hf_lower = hf.replace(" ", "").lower()
            for j in range(len(inventory_header)):
                h = inventory_header[j]
                h_lower = h.replace(" ", "").lower()
                if hf_lower == h_lower:
                    location_inve.append(excel_column_name(j+1))
                    location_inve_index.append(j)
                    break
        
        for i in range(len(location_sale_index)):
            if i != location_sale_index[i]:
                sale_flag = False

        for i in range(len(location_inve_index)):
            if i != location_inve_index[i]:
                inve_flag = False
        data = {}
        dealer_name = dealer_config[f"Dealer{index}"]["DealerName"]
        data["DealerName"] = dealer_name
        data.setdefault("SaleFile", {})["Headerindex"] = "Default" \
            if sale_flag else location_sale
        data.setdefault("InventoryFile", {})["Headerindex"] = "Default" \
            if inve_flag else location_inve
        write_data[f"Dealer{index}"] = data

    with open(ConfigInfo.HeaderChangePath, "w", encoding = "UTF-8")as file:
        json.dump(write_data, file, ensure_ascii=False, indent=2)

# 撰寫sub_record.json檔案
def SubRecordJson(mode, data):
    if mode == "Start": # data = None
        dealer_config = DealerConf()
        dealer_list = dealer_config["DealerList"]
        data = {"StartIndex" : None}
        for i in range(len(dealer_list)):
            index = i + 1
            data[f"Dealer{index}"] = {"SaleFile" : None, "InventoryFile" : None, "Mail2":0, "Mail3":0, "Mail4":0}
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(data, file_json, ensure_ascii = False, indent = 4)
        return "成功建立檔案繳交狀態json檔。"
    
    elif mode == "ReadIndex": # data = None
        running_data = read_json(ConfigInfo.SubRecordPath)
        start_index = running_data["StartIndex"]
        return start_index
    
    elif mode == "WriteIndex": # data = int
        running_data = read_json(ConfigInfo.SubRecordPath)
        running_data["StartIndex"] = data
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(running_data, file_json, ensure_ascii = False, indent = 4)
        return f"更新起始索引為： {data}."
    
    elif mode == "WriteFileStatus": # data = {"Dealer1":{"SaleFile":Ture}}
        sale = "SaleFile"
        inventory = "InventoryFile"
        mail2 = "Mail2"
        mail3 = "Mail3"
        mail4 = "Mail4"
        with open(ConfigInfo.SubRecordPath, "r", encoding = "UTF-8") as file_json:
            running_data = json.load(file_json)
        for i in data:
            if sale in data[i]:
                running_data[i][sale] = data[i][sale]
            elif inventory in data[i]:
                running_data[i][inventory] = data[i][inventory]
            elif mail2 in data[i]:
                running_data[i][mail2] = data[i][mail2]
            elif mail3 in data[i]:
                running_data[i][mail3] = data[i][mail3]
            elif mail4 in data[i]:
                running_data[i][mail4] = data[i][mail4]
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(running_data, file_json, ensure_ascii = False, indent = 4)
        return f"已更新檔案繳交參數。 Data = {data}."
    
    elif mode == "Read": # data = None
        with open(ConfigInfo.SubRecordPath, "r", encoding = "UTF-8") as file_json:
            running_data = json.load(file_json)
        return running_data
    else: # data = None
        return False

# 更新系統使用者名稱至system.json
def WrightWinUser(user_name):
    data = Config()
    data["App"]["WinUser"] = user_name
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"更新使用者名稱：{user_name}。"
    return msg

# 更新OneDrive路徑至system.json
def WrightOneDrivePath(file_path):
    data = Config()
    data["App"]["OneDrivePath"] = file_path
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"OneDrive目錄更新，目錄：{file_path}。"
    return msg

def WrightFileJson(data):
    with open(ConfigInfo.FileConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"File.json BA資訊更新，{data['FileInfo']}。"
    return msg

if __name__ == "__main__":
    # DealerJson()
    # HeaderChange()
    # mode = "WriteFileStatus"
    # data = {"Dealer1":{"SaleFile":True}}
    # SubRecordJson(mode, data)
    aa = Config()
    print(aa)