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

#RawData
RawDataFileName = GlobalConfig["RawDataTable"]["FileName"]
RawDataFileName = RawDataFileName.replace("{Year}", str(Year))
RawDataSheetName = GlobalConfig["RawDataTable"]["SheetName"]
RawDataSheetName = RawDataSheetName.replace("{Month}", str(Month))
RawDataHeader = GlobalConfig["RawDataTable"]["Header"]

#Summary
SummaryReport = GlobalConfig["SummaryTable"]["FileName"]
SummaryReport = SummaryReport.replace("{Year}", str(Year)).replace("{Month}", str(Month))
SummarySheetName = GlobalConfig["SummaryTable"]["SheetName"]
SummarySheetName = SummarySheetName.replace("{Month}", str(Month))
SummaryHeader = GlobalConfig["SummaryTable"]["Header"]

# NotSubmission
NotSubmissionFileName = GlobalConfig["NotSubmissionTable"]["FileName"]
NotSubmissionFileName = NotSubmissionFileName.replace("{Year}", str(Year))
NotSubmissionSheetName = GlobalConfig["NotSubmissionTable"]["SheetName"]
NotSubmissionSheetName = NotSubmissionSheetName.replace("{Month}", str(Month))
NotSubmissionHeader = GlobalConfig["NotSubmissionTable"]["Header"]

ReportDir = GlobalConfig["Default"]["ReportDir"]
DealerDir = GlobalConfig["Default"]["DealerFolderName"]
FolderName = GlobalConfig["App"]["Name"] if GlobalConfig["App"]["Name"] \
    else GlobalConfig["Default"]["Name"]
RootDir = GlobalConfig["App"]["DataPath"] if GlobalConfig["App"]["DataPath"] \
    else GlobalConfig["Default"]["DataPath"]
DealerList = DealerConfig["DealerList"]

Dir = os.path.join(RootDir, FolderName, DealerDir, ReportDir)
RawDataFilePath = os.path.join(Dir, RawDataFileName)
SummaryFilePath = os.path.join(Dir, SummaryReport)
NotSubmissionPath = os.path.join(Dir, NotSubmissionFileName)
RawDataFixedColumns = [chr(i % 26 + 65) for i in range(len(RawDataHeader))]
SummaryFixedColumns = [chr(i % 26 + 65) for i in range(len(SummaryHeader))]
NotSubmissionFixedColumns = [chr(i % 26 + 65) for i in range(len(NotSubmissionHeader))]
ExcelStyle = Alignment(horizontal = "center", vertical = "center")

# 設定Excel的樣式
def set_cell_styles(ws, fixed_columns, width):
    for col in fixed_columns:
        ws.column_dimensions[col].width = width
        for row in range(1, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = ExcelStyle

    for i in range (2, len(DealerList)*2+1,2):
        ws.merge_cells(f"A{i}:A{i + 1}")
        ws.merge_cells(f"B{i}:B{i + 1}")

# 建立固定表頭
def make_part1_format(wb, fixed_columns, header, file_path, sheet_name, width):
    ws = wb.create_sheet(title = sheet_name)
    for col, data in zip(fixed_columns, header):
        ws[f"{col}1"] = data
    
    set_cell_styles(ws, fixed_columns, width)
    wb.save(file_path)

# 寫入經銷商資料
def write_dealer_info(wb, sheet_name, file_path, fixed_columns, width):
    ws = wb[sheet_name]
    dealer_dt_list = ["Sale", "Inventory"]
    dealer_name_list, sale_PC_list, inventory_PC_list = [], [], []

    for i in range(len(DealerList)):
        key1 = f"Dealer{i+1}"
        dealer_name = DealerConfig[key1]["DealerName"]
        sale_PC = DealerConfig[key1]["SaleFile"]["PaymentCycle"]
        inventory_PC = DealerConfig[key1]["InventoryFile"]["PaymentCycle"]
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

    set_cell_styles(ws, fixed_columns, width)
    wb.save(file_path)
    wb.close()

# 創建 RAW Data excel 或 繳交總表
def make_record_tamplates(mode):
    if mode == "report":
        file_path = RawDataFilePath
        file_name = RawDataFileName
        sheet_name = RawDataSheetName
        fixed_columns = RawDataFixedColumns
        header = RawDataHeader
        width = 15
    elif mode == "summary":
        file_path = SummaryFilePath
        file_name = SummaryReport
        sheet_name = SummarySheetName
        fixed_columns = SummaryFixedColumns
        header = SummaryHeader
        width = 22.5

    if os.path.exists(Dir):
        try:
            wb = load_workbook(file_path)
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                first_column_data = []
                for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                    first_column_data.append(row[0])
                dealer_id_in_data = set(first_column_data)
                dealer_list = set(DealerList)
                result = dealer_list.issubset(dealer_id_in_data)
                if result:
                    msg = f"{file_name} 檔案中 {sheet_name} 工作表已存在"
                    WSysLog("1", "MakeRecordTamplates", msg)
                else:
                    write_dealer_info(wb, sheet_name, file_path, fixed_columns, width)
                    msg = f"{file_name} 檔案中 {sheet_name} 工作表更新經銷商資訊"
                    WSysLog("1", "MakeRecordTamplates", msg)
            else:
                make_part1_format(wb, fixed_columns, header, file_path, sheet_name, width)
                write_dealer_info(wb, sheet_name, file_path, fixed_columns, width)
                msg = f"{file_name} 檔案中 {sheet_name} 工作表已建立"
                WSysLog("1", "MakeRecordTamplates", msg)
            return True
        
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            make_part1_format(wb, fixed_columns, header, file_path, sheet_name, width)
            write_dealer_info(wb, sheet_name, file_path, fixed_columns, width)
            msg = f"{file_name} 檔案已建立"
            WSysLog("1", "MakeRecordTamplates", msg)
            return True
    else:
        msg = f"{Dir}目錄不存在"
        WSysLog("3", "MakeRecordTamplates", msg)
        return False

# 寫入RawData
def WriteRawData(new_data):
    result = make_record_tamplates("report")
    if result:
        wb = load_workbook(RawDataFilePath)
        ws = wb[RawDataSheetName]
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
            msg = f"{RawDataFileName} 檔案，{RawDataSheetName} 工作表，新增資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
        else:
            col_str = get_column_letter(col_idx)
            msg = f"{RawDataFileName} 檔案，{RawDataSheetName} 工作表，{column_name} 更新資料：{new_data}"
            WSysLog("1", "WriteRawData", msg)
        
        # 寫入當天資料
        for i, value in enumerate(new_data, start = 2):
            ws.cell(row = i, column = col_idx, value = value)
            ws[f"{col_str}{i}"].alignment = ExcelStyle
        
        # 寫入當天更新筆數欄位
        update_num = [item.split("/")[0] for item in new_data]
        for i, value in enumerate(update_num, start = 2):
            ws.cell(row = i, column = 5, value = value)
        wb.save(RawDataFilePath)

    else:
        msg = f"{RawDataFileName} 檔案 {RawDataSheetName} 新增資料失敗"
        WSysLog("2", "WriteRawData", msg)

# 繳交總表(每月一份檔案), not_finish
def WriteSummaryData():
    result = make_record_tamplates("summary")
    if result:
        wb = load_workbook(SummaryFilePath)
        ws = wb[SummarySheetName]

    else:
        print("")

# NotSubmission
# 設定缺繳、補繳工作表欄寬
def set_not_sub_styles(ws):
    column_widths = [15, 15, 15, 15, 35, 15, 10, 10, 15, 15]
    for i, width in enumerate(column_widths):
        ws.column_dimensions[NotSubmissionFixedColumns[i]].width = width

# 產生缺繳、補繳工作表表頭
def make_not_sub_header(wb):
    ws = wb.create_sheet(title = NotSubmissionSheetName)
    for col, data in zip(NotSubmissionFixedColumns, NotSubmissionHeader):
        ws[f"{col}1"] = data
    set_not_sub_styles(ws)
    wb.save(NotSubmissionPath)

# 產生補繳、缺繳工作表
def make_not_sub_table():
    if os.path.exists(Dir):
        try:
            wb = load_workbook(NotSubmissionPath)
            if NotSubmissionSheetName in wb.sheetnames:
                msg = f"{NotSubmissionFileName} 檔案中 {NotSubmissionSheetName} 工作表已存在"
                WSysLog("1", "MakeNotSubTable", msg)
            else:
                make_not_sub_header(wb)
                msg = f"{NotSubmissionFileName} 檔案中 {NotSubmissionSheetName} 工作表已建立"
                WSysLog("1", "MakeNotSubTable", msg)
            return True
        except FileNotFoundError:
            wb = Workbook()
            wb.remove(wb.active)
            make_not_sub_header(wb)
            msg = f"{NotSubmissionFileName} 檔案已建立"
            WSysLog("1", "MakeNotSubTable", msg)
            return True
    else:
        msg = f"{Dir}目錄不存在"
        WSysLog("3", "MakeNotSubTable", msg)
        return False

# 寫入NotSubmission，["Dealer ID", "Dealer Name", "DataType", "檔案繳交週期", "缺繳檔案名稱", "應繳時間", "是否繳交", "檢查結果", "補繳時間", "補繳檢查結果"]
# 資料一row一row寫
def WriteNotSubmission(write_data):
    result = make_not_sub_table()
    if result:
        wb = load_workbook(NotSubmissionPath)
        ws = wb[NotSubmissionSheetName]
        not_sub_file_name = [cell.value for cell in ws['E']]
        index = None
        for row in range(len(not_sub_file_name)):
            if write_data[4] == not_sub_file_name[row]:
                index = row + 1
                break

        if index is not None:
            write_data = write_data[-2:]
            for i in range(len(write_data)):
                col = chr((8 + i) % 26 + 65)
                ws[f"{col}{index}"] = write_data[i]
            msg = f"{NotSubmissionFileName} 檔案 {NotSubmissionSheetName} 工作表中，row{index} 資料更新，補繳時間：{write_data[0]}，補繳檢查結果：{write_data[1]}"
            WSysLog("1", "WriteNotSubmission", msg)
        else:
            index = ws.max_row + 1
            for col, data in zip(NotSubmissionFixedColumns, write_data):
                ws[f"{col}{index}"] = data
            msg = f"{NotSubmissionFileName} 檔案 {NotSubmissionSheetName} 工作表中，row{index} 新增資料，{write_data}"
            WSysLog("1", "WriteNotSubmission", msg)
        wb.save(NotSubmissionPath)
    else:
        msg = f"{NotSubmissionFileName} 檔案 {NotSubmissionSheetName} 新增資料失敗"
        WSysLog("2", "WriteNotSubmission", msg)

if __name__ == "__main__":
    make_record_tamplates("summary")
    make_record_tamplates("report")
    # input_data = ["000"] * 8
    # input_data.append("001")
    # input_data.append("002")
    # print(input_data)
    # WriteNotSubmission(input_data)