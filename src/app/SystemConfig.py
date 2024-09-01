# -*- coding: utf-8 -*-

'''
檔案說明：讀取、變更 ./src/config 底下的 json 檔案
Writer：Qian
'''

# 標準庫
import json

# 自定義函數
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

# 撰寫sub_record.json檔案
def SubRecordJson(mode, data):
    if mode == "Start": # data = None
        dealer_config = DealerConf()
        dealer_list = dealer_config["DealerList"]
        data = {"SubStartIndex" : None, "NotSubStartIndex" : None, "ChangeDic" : None}
        for i in range(len(dealer_list)):
            index = i + 1
            data[f"Dealer{index}"] = {"SaleFile" : None, "InventoryFile" : None, "Mail2":{}, "Mail3":{}, "Mail4":{}}
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(data, file_json, ensure_ascii = False, indent = 4)
        return "成功建立檔案繳交狀態json檔。"

    elif mode == "ReadSubStartIndex": # data = None
        running_data = read_json(ConfigInfo.SubRecordPath)
        start_index = running_data["SubStartIndex"]
        return start_index

    elif mode == "WriteSubStartIndex": # data = int
        running_data = read_json(ConfigInfo.SubRecordPath)
        running_data["SubStartIndex"] = data
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(running_data, file_json, ensure_ascii = False, indent = 4)
        return f"更新起始索引為： {data}."

    elif mode == "ReadNotSubStartIndex": # data = None
        running_data = read_json(ConfigInfo.SubRecordPath)
        start_index = running_data["NotSubStartIndex"]
        return start_index

    elif mode == "WriteNotSubStartIndex": # data = int
        running_data = read_json(ConfigInfo.SubRecordPath)
        running_data["NotSubStartIndex"] = data
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
                file_name = [j for j in data[i][mail2]]
                if file_name:
                    running_data[i][mail2][file_name[0]] = data[i][mail2][file_name[0]]
                else:
                    running_data[i][mail2] = data[i][mail2]
            elif mail3 in data[i]:
                file_name = [j for j in data[i][mail3]]
                if file_name:
                    running_data[i][mail3][file_name[0]] = data[i][mail3][file_name[0]]
                else:
                    running_data[i][mail3] = data[i][mail3]
            elif mail4 in data[i]:
                file_name = [j for j in data[i][mail4]]
                if file_name:
                    running_data[i][mail4][file_name[0]] = data[i][mail4][file_name[0]]
                else:
                    running_data[i][mail4] = data[i][mail4]

        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(running_data, file_json, ensure_ascii = False, indent = 4)
        return f"已更新檔案繳交參數。 Data = {data}."

    elif mode == "WriteChangeDic": # data = {id:filename}
        running_data = read_json(ConfigInfo.SubRecordPath)
        running_data["ChangeDic"] = data
        with open(ConfigInfo.SubRecordPath, "w", encoding = "UTF-8") as file_json:
            json.dump(running_data, file_json, ensure_ascii = False, indent = 4)
        return f"已更新ChangeDic = {data}."

    elif mode == "ReadChangeDic":
        with open(ConfigInfo.SubRecordPath, "r", encoding = "UTF-8") as file_json:
            running_data = json.load(file_json)
        return running_data["ChangeDic"]

    elif mode == "Read": # data = None
        with open(ConfigInfo.SubRecordPath, "r", encoding = "UTF-8") as file_json:
            running_data = json.load(file_json)
        return running_data
    else: # data = None
        return False

# 更新系統使用者名稱至system.json
def WriteWinUser(user_name):
    data = Config()
    data["App"]["WinUser"] = user_name
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"更新使用者名稱：{user_name}。"
    return msg

# 更新OneDrive路徑至system.json
def WriteOneDrivePath(file_path):
    data = Config()
    data["App"]["OneDrivePath"] = file_path
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"OneDrive目錄更新，目錄：{file_path}。"
    return msg

# 更新工作日資訊至system.json
def WrtieWorkDay(work_day):
    data = Config()
    data["Default"]["WorkDay"] = work_day
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"更新WorkDay參數： {work_day}。"
    return msg

def WriteWorkDayCounter(num):
    data = Config()
    data["Default"]["WorkDayCounter"] = num
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"更新WorkDayCounter參數： {num}。"
    return msg

def WriteMonthlySubFlag(status):
    data = Config()
    data["Default"]["MonthlySubFlag"] = status
    with open(ConfigInfo.ConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"更新MonthlySubFlag參數： {status}。"
    return msg

# 將檔案更新時間寫入 files.json
def WriteFileJson(data):
    with open(ConfigInfo.FileConfigPath, "w", encoding = "UTF-8") as file:
        json.dump(data, file, ensure_ascii = False, indent = 4)
    msg = f"File.json BA資訊更新，{data['FileInfo']}。"
    return msg

# 寫入Dealer.json
def WriteDealerJson(mode, data):
    # 將新的 DealerList 寫入 dealer.json
    if mode == "DealerList":
        dealer_list = data
        data = DealerConf()
        data["DealerList"] = dealer_list
        with open(ConfigInfo.DealerPath, "w", encoding = "UTF-8") as file:
            json.dump(data, file, ensure_ascii = False, indent = 4)
        return True
    # Dealer.json寫入新的新銷商資訊
    elif mode == "DealerInfo":
        with open(ConfigInfo.DealerPath, "w", encoding = "UTF-8") as file:
            json.dump(data, file, ensure_ascii = False, indent = 4)
        return True

if __name__ == "__main__":
    # DealerJson()
    # HeaderChange()
    # mode = "WriteFileStatus"
    # data = {"Dealer1":{"SaleFile":True}}
    Mode = "Start"
    Data = None
    SubRecordJson(Mode, Data)
    # aa = Config()
    # print(aa)