# -*- coding: utf-8 -*-

'''
檔案說明：讀取、變更System.json檔案
Writer：Qian
'''

import os
import json
import pandas as pd

ConfigDir = "src/config"
SystemConfig = "system.json"
MappingConfig = "mapping_rule.json"
CheckConfig = "check_rule.json"

ConfigPath = os.path.join(ConfigDir, SystemConfig)
MappingPath = os.path.join(ConfigDir, MappingConfig)
CheckPath = os.path.join(ConfigDir,CheckConfig)

def Config():
    with open(ConfigPath,"r",encoding="utf-8") as file:
        config = json.load(file)
    return config

def MappingRule():
    with open(MappingPath, "r", encoding = "UTF-8") as file:
        mapping_config = json.load(file)
    return mapping_config

def CheckRule():
    with open(CheckPath, "r", encoding = "UTF-8") as file:
        check_config = json.load(file)
    return check_config

def WriteRuleJson(file_path, sheet, data_name, output_name):
    df = pd.read_excel(file_path,sheet_name=sheet)
    data_list = df.to_dict("records")
    final_json = {
        data_name: data_list
    }
    output_path = os.path.join("./src/config", output_name)
    with open(output_path, "w",encoding="UTF-8") as f:
        json.dump(final_json, f, ensure_ascii=False, indent=2)

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

def MakeMappingRuleJson():
    DataName = ["Sale","Inventory"]
    FilePath = "./docs/data_format/DataFormatFromBD.xlsx"
    Sheet = "FileMappingRule"
    Output = "MappingRuleOutput.json"
    for i in DataName:
        SheetName = i + Sheet
        OutputName = i + Output
        WriteRuleJson(FilePath, SheetName, i, OutputName)
    File1Name = str(DataName[0]) + Output
    File2Name = str(DataName[1]) + Output
    marge_json_files(ConfigDir, File1Name, File2Name, MappingConfig)

if __name__ == "__main__":
    MakeMappingRuleJson()