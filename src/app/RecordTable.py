# -*- coding: utf-8 -*-

'''
檔案說明：撰寫繳交紀錄表
Writer：Qian
'''

import os
import pandas as pd
from Log import WSysLog
from datetime import datetime
from openpyxl.styles import Alignment
from SystemConfig import Config, DealerConf
from openpyxl import Workbook, load_workbook

GlobalConfig = Config()
DealerConfig = DealerConf()
CurrentDate = datetime.now()
Day, Month, Year = CurrentDate.day, CurrentDate.month, CurrentDate.year

ReportFileName = f"{Year}_RawData.xlsx"
SheetName = f"{Month}月"
RawDataHeader = ["DealerID","DealerName","DataType","檔案繳交週期","當日更新筆數"]

ReportDir = GlobalConfig["Default"]["ReportDir"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]

Dir = os.path.join(RootDir, FolderName, DealerDir, ReportDir)
ReportFilePath = os.path.join(Dir, ReportFileName)

# RAW Data
# 創建左側 A ~ E 欄固定值
def make_part1_format(wb):
    ws = wb.create_sheet(title = SheetName)
    fixed_columns = [chr(i % 26 + 65) for i in range(5)]
    for col, data in zip(fixed_columns, RawDataHeader):
        ws[f"{col}1"] = data
    set_cell_styles(ws)
    wb.save(ReportFilePath)

def set_cell_styles(ws):
    fixed_columns = [chr(i % 26 + 65) for i in range(5)]
    style = Alignment(horizontal = "center", vertical = "center")
    for col in fixed_columns:
        ws.column_dimensions[col].width = 15
        for row in range(1, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = style

# 寫入經銷商資料
def write_dealer_info():
    wb = load_workbook(ReportFilePath)
    ws = wb[SheetName]
    dealer_dt_list = ["Sale", "Inventory"]
    dealer_name_list, dealer_PC_list = [], []
    dealer_list = DealerConfig["DealerList"]

    fixed_columns = [chr(i % 26 + 65) for i in range(5)]
    for col, data in zip(fixed_columns, RawDataHeader):
        ws[f"{col}1"] = data

    for i in range(len(dealer_list)):
        key1 = f"Dealer{i+1}"
        dealer_name = DealerConfig[key1]["DealerName"]
        sale_PC = DealerConfig[key1]["SaleFile"]["PaymentCycle"]
        inventory_PC = DealerConfig[key1]["InventoryFile"]["PaymentCycle"]
        dealer_name_list.append(dealer_name)
        dealer_PC_list.append(sale_PC)
        dealer_PC_list.append(inventory_PC)

    for i in range(2, len(dealer_list)*2+1,2):
        ws.merge_cells(f"A{i}:A{i + 1}")
        ws.merge_cells(f"B{i}:B{i + 1}")
        
        ws[f"A{i}"] = dealer_list[int(i / 2 - 1)]
        ws[f"B{i}"] = dealer_name_list[int(i / 2 - 1)]
        ws[f"C{i}"] = dealer_dt_list[int(i % 2)]
        ws[f"C{i + 1}"] = dealer_dt_list[int((i+1) % 2)]
        ws[f"D{i}"] = dealer_PC_list[int(i / 2 - 1)]
        ws[f"D{i+1}"] = dealer_PC_list[int(i / 2)]
    
    set_cell_styles(ws)
    wb.save(ReportFilePath)
    wb.close()

# 創建 RAW Data excel 及 工作表
def make_record_tamplates():
    if os.path.exists(Dir):
        try:
            wb = load_workbook(ReportFilePath)
            if SheetName in wb.sheetnames:
                msg = f"{ReportFileName} 檔案中 {SheetName} 工作表已存在"
                WSysLog("1", "MakeRecordTamplates", msg)
            else:
                make_part1_format(wb)
                msg = f"{ReportFileName} 檔案中 {SheetName} 工作表已建立"
                WSysLog("1", "MakeRecordTamplates", msg)
            return True
        
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            make_part1_format(wb)
            write_dealer_info()
            msg = f"{ReportFileName} 檔案已建立"
            WSysLog("1", "MakeRecordTamplates", msg)
            return True
    else:
        msg = f"{Dir}目錄不存在"
        WSysLog("3", "MakeRecordTamplates", msg)
        return False

# 新資料寫入
def WriteRawData(new_data):
    result = make_record_tamplates()
    if result:
        df = pd.read_excel(ReportFilePath, sheet_name = SheetName)
        column_name = f"{Month}月{Day}日"
        file_header = df.columns.tolist()
        df[column_name] = new_data
        df.to_excel(ReportFilePath, sheet_name = SheetName, index = False)
        if column_name in file_header:
            msg = f"{ReportFileName} 檔案，{SheetName} 工作表，{column_name} 更新資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
        else:
            msg = f"{ReportFileName} 檔案，{SheetName} 工作表，新增資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
    else:
        msg = f"{ReportFileName} 檔案 {SheetName} 新增資料失敗"
        WSysLog("2", "WriteRawData", msg)

if __name__ == "__main__":
    # wb = load_workbook(ReportFilePath)
    # write_dealer_info(wb)
    NewData = ["00","11"]
    WriteRawData(NewData)