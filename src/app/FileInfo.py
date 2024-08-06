import os
from datetime import datetime
from Log import WSysLog
from Config import AppConfig
from SystemConfig import WrightFileJson
from openpyxl import Workbook, load_workbook

Config = AppConfig()

# 產生對應的 Excel Column 名稱(一次多個)
def get_excel_colmun_name(file_header):
    return [chr(i % 26+ 65) for i in range(len(file_header))]

# 根據表頭搜尋出 Excel 的 Column 名稱(一次一個)
def search_column_name(file_header, col_name):
    for index in range(len(file_header)):
        if file_header[index] == col_name:
            break
    return chr(index % 26 + 65)

# 設定表格樣式
def excel_style(ws, column_width, fixed_columns):
    for i, width in enumerate(column_width):
        ws.column_dimensions[fixed_columns[i]].width = width
    for col in fixed_columns:
        for row in range(1, ws.max_row + 1):
            ws[f"{col}{row}"].alignment = Config.ExcelStyle

# 從使用者資訊中取出BA名單
def get_BA():
    user = Config.UserConfig
    ba_list = []
    for user_id in user:
        if user[user_id]["Group"] == "BD_BA":
            BA_id_part = user[user_id]["Description"].split(" ")
            BA_id = BA_id_part[1]
            BA_name = user[user_id]["Name"]
            BA_Dealer_id = user[user_id]["ResponsibleDealerID"]
            ba_list.append({"BA_ID":BA_id, "BA_Name":BA_name, "BA_Dealer_ID" : BA_Dealer_id})
    return ba_list

# 依照檔案格式建立空白內容之檔案
def make_format_file(folder_path, file_name, sheet_name, file_header, column_width):
    try:
        file_path = os.path.join(folder_path, file_name)
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet(title = sheet_name)
        fixed_columns = get_excel_colmun_name(file_header)
        for col, data in zip(fixed_columns, file_header):
            ws[f"{col}1"] = data
        excel_style(ws, column_width, fixed_columns)
        wb.save(file_path)
        msg = f"成功於 {folder_path} 目錄中建立空白 {file_name} 檔案。"
        WSysLog("1", "MakeFormatFile", msg)
    except Exception as e:
        msg = f"於 {folder_path} 建立空白 {file_name} 檔案時發生錯誤： {e}。"
        WSysLog("3", "MakeFormatFile", msg)

# 將BA負責的經銷商ID寫入DealerInfo
def wright_dealer_info_in_file(file_path, sheet_name, file_header, BA_dealer_id):
    try:
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
        position_list = ["業務窗口","專案聯絡人(IT人員 or 資料管理人員","財務窗口"]
        content_col = ["Position", "Name", "Mail", "Ex"]
        for row in range(2, (len(BA_dealer_id))*3+1, 3):
            index = int((row - 2) / 3)
            col =  search_column_name(file_header, "Dealer ID")
            ws[f"{col}{row}"] = BA_dealer_id[index]
            
            for j in range(3):
                col = search_column_name(file_header, "Position")
                ws[f"{col}{row + j}"] = position_list[j]
                for col_name in content_col:
                    col = search_column_name(file_header, col_name)
                    ws[f"{col}{row + j}"].alignment = Config.ExcelStyle
            
            for col_name in file_header:
                if col_name in content_col:
                    continue
                col = search_column_name(file_header, col_name)
                ws.merge_cells(f"{col}{row}:{col}{row + j}")
                ws[f"{col}{row}"].alignment = Config.ExcelStyle
        wb.save(file_path)
        msg = f"成功寫入經銷商資訊至 {file_path}。"
        WSysLog("1", "WrightDealerInfoInFile", msg)
    except Exception as e:
        msg = f"寫入經銷商資訊發生錯誤： {e}。"
        WSysLog("3", "WrightDealerInfoInFile", msg)

# 檢查BA目錄底下檔案
def CheckBAFolderFiles():
    BA_list = get_BA()
    flag = True
    master_file = Config.MasterFileName
    dealer_info_file = Config.DealerInfoFileName
    kalist_file = Config.KAListFileName
    file_time_dic = {}
    for ba in BA_list:
        file_time = {}
        BA_id = ba["BA_ID"]
        BA_name = ba["BA_Name"]
        BA_dealer_id = ba["BA_Dealer_ID"]
        new_ba_id = BA_id[:2] + "_" + BA_id[2:]
        ba_folder_name = new_ba_id + "_" + BA_name
        folder_path = os.path.join(Config.BAFolderPath, ba_folder_name)

        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        master_file_path = os.path.join(folder_path, master_file)
        dealer_info_file_path = os.path.join(folder_path, dealer_info_file)

        # 目標目錄找不到 masterfile 時須建立一份空的 masterfile
        if not os.path.exists(master_file_path):
            flag = False
            msg = f"{ba_folder_name}目錄下缺少{master_file}檔案，系統將建立空白檔案。"
            WSysLog("2", "CheckBAFolderFiles", msg)
            sheet_name = Config.MasterFileSheetName
            file_header = Config.MasterFileHeader
            column_width = Config.MasterFileColumnWidth
            make_format_file(folder_path, master_file, sheet_name, file_header, column_width)
            file_time["MasterFile"] = None
        else:
            file_time["MasterFile"] = datetime.fromtimestamp(os.path.getatime(master_file_path))

        # 目標目錄找不到 dealerinfo 時須建立一份空的 dealerinfo
        if not os.path.exists(dealer_info_file_path):
            flag = False
            msg = f"{ba_folder_name}目錄下缺少{dealer_info_file}檔案，系統將建立空白檔案。"
            WSysLog("2", "CheckBAFolderFiles", msg)
            sheet_name = Config.DealerInfoFileSheetName
            file_header = Config.DealerInfoFileHeader
            column_width = Config.DealerInfoFileColumnWidth
            make_format_file(folder_path, dealer_info_file, sheet_name, file_header, column_width)
            wright_dealer_info_in_file(dealer_info_file_path, sheet_name, file_header, BA_dealer_id)
            file_time["DealerInfo"] = None
        else:
            file_time["DealerInfo"] = datetime.fromtimestamp(os.path.getatime(dealer_info_file_path))

        # KAList
        for ka_dealer in Config.KADealerList:
            if ka_dealer in BA_dealer_id:
                ka_list_file_path = os.path.join(folder_path, kalist_file)
                
                # 目標目錄找不到kalist時須建立一份空的kalist
                if not os.path.exists(ka_list_file_path):
                    flag = False
                    msg = f"{ba_folder_name}目錄下缺少{kalist_file}檔案，系統將建立空白檔案。"
                    WSysLog("2", "CheckBAFolderFiles", msg)
                    sheet_name = Config.KAListFileSheetName
                    file_header = Config.KAListFileHeader
                    for i in range(len(Config.DealerList)):
                        if Config.DealerList[i] == ka_dealer:
                            index = i + 1
                            break
                    dealer_name = Config.DealerConfig[f"Dealer{index}"]["DealerName"]
                    sheet_name = ka_dealer + "_" + sheet_name
                    file_header[0] = dealer_name + file_header[0]
                    column_width = Config.KAListFileColumnWidth
                    make_format_file(folder_path, kalist_file, sheet_name, file_header, column_width)
                    file_time["KAList"] = None
                else:
                    file_time["KAList"] = datetime.fromtimestamp(os.path.getatime(ka_list_file_path))

        file_time_dic[BA_id] = file_time
    
    if flag:
        msg = "BA目錄底下應有檔案都存在。" 
        WSysLog("1", "CheckBAFolderFiles", msg)
        return file_time_dic
    else:
        return None
    
# 依據BA人數調整files.json格式
def check_ba():
    flag = True
    BA_list = get_BA()
    json_data = Config.FileConfig

    for ba in BA_list:
        BA_id = ba["BA_ID"]
        if BA_id not in json_data["FileInfo"]:
            flag = False
            json_data["FileInfo"][BA_id] = {"MasterFile":None, "KAList":None, "DealerInfo":None}

    if flag:
        msg = "files.json資料無異動。"
        WSysLog("1", "CheckBA", msg)
    else:
        msg = WrightFileJson(json_data)
        WSysLog("1", "CheckBA", msg)

# 檢查檔案資訊主程式
def CheckFileInfo():
    result = CheckBAFolderFiles()
    if result is not None:
        check_ba()
        BA_list = get_BA()


if __name__ == "__main__":
    CheckBAFolderFiles()