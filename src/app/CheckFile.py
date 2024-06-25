# -*- coding: utf-8 -*-

'''
檔案說明：確認檔案繳交時間，
檢查檔案副檔名、表頭格式及內容
Writer:Qian
'''
# import os
import pandas as pd
# import openpyxl
from SystemConfig import Config, CheckRule

GlobalConfig = Config()
CheckConfig = CheckRule()
SF_MustHave = CheckConfig["SaleFile"]["MustHave"]
IF_MustHave = CheckConfig["InventoryFile"]["MustHave"]

# 記錄檔案為正常繳交、缺交
# def FileUploadTime():


def CheckSaleFile(FilePath):
    df = pd.read_excel(FilePath)
    print(df)
# with

if __name__ == "__main__":
    FilePath = "./datas/test.xlsx"
    CheckSaleFile(FilePath=FilePath)