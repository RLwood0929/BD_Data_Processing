# -*- coding: utf-8 -*-

'''
檔案說明：讀取、變更 ./src/config 底下的 json 檔案
Writer：Qian
'''

import os
import json
import pandas as pd

ConfigDir = "src/config"
SystemConfig = "system.json"
MappingConfig = "mapping_rule.json"
CheckConfig = "check_rule.json"
DealerConfig = "dealer.json"

ConfigPath = os.path.join(ConfigDir, SystemConfig)
MappingPath = os.path.join(ConfigDir, MappingConfig)
CheckPath = os.path.join(ConfigDir, CheckConfig)
DealerPath = os.path.join(ConfigDir, DealerConfig)

# 讀取 system.json
def Config():
    with open(ConfigPath,"r",encoding="utf-8") as file:
        config = json.load(file)
    return config

# 讀取 mapping_rule.json
def MappingRule():
    with open(MappingPath, "r", encoding = "UTF-8") as file:
        mapping_config = json.load(file)
    return mapping_config

# 讀取 check_rule.json
def CheckRule():
    with open(CheckPath, "r", encoding = "UTF-8") as file:
        check_config = json.load(file)
    return check_config

# 讀取 dealer.json
def DealerConf():
    with open(DealerPath, "r", encoding = "UTF-8") as file:
        dealer_config = json.load(file)
    return dealer_config

#將 excel 內容轉變為json
def write_rule_json(file_path, sheet, data_name, output_name):
    df = pd.read_excel(file_path,sheet_name=sheet)
    data_list = df.to_dict("records")
    final_json = {
        data_name: data_list
    }
    output_path = os.path.join("./src/config", output_name)
    with open(output_path, "w",encoding="UTF-8") as f:
        json.dump(final_json, f, ensure_ascii=False, indent=2)

# 合併json檔案
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

# 將 maping rule 的excel轉變為json
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
    marge_json_files(ConfigDir, File1Name, File2Name, MappingConfig)

# 將經銷商資訊轉成 json
def DealerJson():
    output_path = os.path.join(ConfigDir, DealerConfig)
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

if __name__ == "__main__":
    DealerJson()