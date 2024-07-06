# -*- coding: utf-8 -*-

'''
檔案說明：檔案進行格式轉換
Writer:Qian
'''

#import os
import pandas as pd
from dateutil import parser
from SystemConfig import Config, MappingRule

GlobalConfig = Config()
MappingConfig = MappingRule()

DataFormatPath = GlobalConfig["DataFormat"]["FilePath"]
MappingSalesSheetName = GlobalConfig["DataFormat"]["MappingSales"]

# SalesInput = ""
# InventoryInput = ""
SalesOutputFormat = GlobalConfig["DataFormat"]["SalesFormatToBD"]
InventoryOutputFormat = GlobalConfig["DataFormat"]["InventoryFormatToBD"]
Mapping = pd.read_excel(DataFormatPath, sheet_name = MappingSalesSheetName)
SalesOutput = pd.read_excel(DataFormatPath, sheet_name = SalesOutputFormat)

DateFormat = MappingConfig["DateFormatChang"]["DateFormat"]

# 解析日期資料來源格式，轉換為規定格式，Rule_1
def DateFormatChang(DateIn):
    try:
        parsed_date = parser.parse(DateIn)
        print(f"原始日期字符串:{DateIn}")
        print(f"解析後的日期:{parsed_date}")
        FormattedDate = parsed_date.strftime(DateFormat)
        print(f"格式 {DateFormat} : {FormattedDate}")
        return FormattedDate
    except ValueError:
        print(f"解析失敗: 無法識別的日期格式 {DateIn}")
        return None

# def SearchUoM(dealerID, prodectID):

# def SearchPrice(dealerID, prodectID, date):

# 固定值填寫，Value
def FixedValue():
    for j in range(5):
        for column in SalesOutput.columns.tolist():
            for _, row in Mapping.iterrows():
                value = row["value"]
                after_conversion = row["After conversion"]
                if pd.notnull(value) and column == after_conversion:
                    SalesOutput.loc[j,column] = value

def MappingData():
    print("")
    OutputFilePath = './datas/test.CSV'  # 替換為你的輸出檔案路徑
    SalesOutput.to_csv(OutputFilePath, index=False)

    print(f"已成功將資料寫入到 {OutputFilePath}")

if __name__ == "__main__":
    datein = [
        "2023-06-19",
        "19/06/2023",
        "June 19, 2023",
        "19 Jun 2023",
        "2023.06.19",
        "2023/06/19"
    ]
    for i in datein:
        DateFormatChang(i)
        print("\n")