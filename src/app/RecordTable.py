# -*- coding: utf-8 -*-

'''
檔案說明：撰寫繳交紀錄表
Writer：Qian
'''

import os
from Log import WSysLog
from datetime import datetime
from openpyxl.styles import Alignment
from SystemConfig import Config, DealerConf
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

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
DealerList = DealerConfig["DealerList"]

Dir = os.path.join(RootDir, FolderName, DealerDir, ReportDir)
ReportFilePath = os.path.join(Dir, ReportFileName)
FixedColumns = [chr(i % 26 + 65) for i in range(5)]
ExcelStyle = Alignment(horizontal = "center", vertical = "center")

# RAW Data
# 創建左側 A ~ E 欄固定值
def make_part1_format(wb):
    ws = wb.create_sheet(title = SheetName)
    for col, data in zip(FixedColumns, RawDataHeader):
        ws[f"{col}1"] = data

    set_cell_styles(ws)
    wb.save(ReportFilePath)

# 設定Excel的樣式
def set_cell_styles(ws):
    for col in FixedColumns:
        ws.column_dimensions[col].width = 15
        for row in range(1, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = ExcelStyle

    for i in range (2, len(DealerList)*2+1,2):
        ws.merge_cells(f"A{i}:A{i + 1}")
        ws.merge_cells(f"B{i}:B{i + 1}")

# 寫入經銷商資料
def write_dealer_info(wb):
    ws = wb[SheetName]
    dealer_dt_list = ["Sale", "Inventory"]
    dealer_name_list, sale_PC_list, inventory_PC_list = [], [], []

    for col, data in zip(FixedColumns, RawDataHeader):
        ws[f"{col}1"] = data

    for i in range(len(DealerList)):
        key1 = f"Dealer{i+1}"
        #print(f"key1:{key1}")
        dealer_name = DealerConfig[key1]["DealerName"]
        sale_PC = DealerConfig[key1]["SaleFile"]["PaymentCycle"]
        #print(f"sale_PC:{sale_PC}")
        inventory_PC = DealerConfig[key1]["InventoryFile"]["PaymentCycle"]
        #print(f"inventory_PC:{inventory_PC}")
        dealer_name_list.append(dealer_name)
        sale_PC_list.append(sale_PC)
        inventory_PC_list.append(inventory_PC)
    
    for i in range(2, len(DealerList)*2+1,2):
        ws[f"A{i}"] = DealerList[int(i / 2 - 1)]
        ws[f"B{i}"] = dealer_name_list[int(i / 2 - 1)]
        ws[f"C{i}"] = dealer_dt_list[int(i % 2)]
        ws[f"C{i + 1}"] = dealer_dt_list[int((i+1) % 2)]
        ws[f"D{i}"] = sale_PC_list[int(i / 2 - 1)]
        ws[f"D{i+1}"] = inventory_PC_list[int(i / 2 - 1)]
    
    set_cell_styles(ws)

    wb.save(ReportFilePath)
    wb.close()

# 創建 RAW Data excel 及 工作表
def make_record_tamplates():
    if os.path.exists(Dir):
        try:
            wb = load_workbook(ReportFilePath)
            if SheetName in wb.sheetnames:
                ws = wb[SheetName]
                first_column_data = []
                for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                    first_column_data.append(row[0])
                dealer_id_in_data = set(first_column_data)
                dealer_list = set(DealerList)
                result = dealer_list.issubset(dealer_id_in_data)
                if result:
                    msg = f"{ReportFileName} 檔案中 {SheetName} 工作表已存在"
                    WSysLog("1", "MakeRecordTamplates", msg)
                else:
                    write_dealer_info(wb)
                    msg = f"{ReportFileName} 檔案中 {SheetName} 工作表更新經銷商資訊"
                    WSysLog("1", "MakeRecordTamplates", msg)
            else:
                make_part1_format(wb)
                write_dealer_info(wb)
                msg = f"{ReportFileName} 檔案中 {SheetName} 工作表已建立"
                WSysLog("1", "MakeRecordTamplates", msg)
            return True
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            make_part1_format(wb)
            write_dealer_info(wb)
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
        wb = load_workbook(ReportFilePath)
        ws = wb[SheetName]
        column_name = f"{Month}月{Day}日"
        col_idx = None

        for col in range(1, ws.max_column + 1):
            if ws.cell(row = 1, column = col).value == column_name:
                col_idx = col
                break
        
        if col_idx is None:
            col_idx = ws.max_column + 1
            ws.cell(row = 1, column = col_idx, value = column_name)
            col_str = get_column_letter(col_idx)
            ws.column_dimensions[col_str].width = 30
            ws[f"{col_str}1"].alignment = ExcelStyle
            msg = f"{ReportFileName} 檔案，{SheetName} 工作表，新增資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
        else:
            col_str = get_column_letter(col_idx)
            msg = f"{ReportFileName} 檔案，{SheetName} 工作表，{column_name} 更新資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
        
        # 寫入當天資料
        for i, value in enumerate(new_data, start = 2):
            ws.cell(row = i, column = col_idx, value = value)
            ws[f"{col_str}{i}"].alignment = ExcelStyle
        
        # 寫入當天更新筆數欄位
        update_num = [item.split("/")[0] for item in new_data]
        for i, value in enumerate(update_num, start = 2):
            ws.cell(row = i, column = 5, value = value)
        wb.save(ReportFilePath)
    else:
        msg = f"{ReportFileName} 檔案 {SheetName} 新增資料失敗"
        WSysLog("2", "WriteRawData", msg)

if __name__ == "__main__":
    NewData = ["66/已繳交/錯誤","11/已繳交/正確","00/未繳交/錯誤","88/已繳交/正確"]
    WriteRawData(NewData)