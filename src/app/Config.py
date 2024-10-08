# -*- coding: utf-8 -*-

"""
檔案說明：存放共用的全域參數
Writer：Qian
"""

# 標準庫
import os
from datetime import datetime
from dotenv import load_dotenv

# 第三方庫
from openpyxl.styles import Alignment

# 自定義函數
from SystemConfig import Config, DealerConf, CheckRule, DealerFormatConf,\
                         MappingRule, MailRule, User, File

# 程式全域參數
class AppConfig:
    def __init__(self):
        # 設定全域系統時間
        self.SystemTime = datetime.now()
        self.Day, self.Month, self.Year = self.SystemTime.day,\
                                          self.SystemTime.month,\
                                          self.SystemTime.year

        # 全域 Excel 樣式
        self.ExcelStyle = Alignment(horizontal = "center", vertical = "center")

        # 從env檔案中取得帳號密碼
        load_dotenv()
        self.EmailSender = os.getenv("Sender")
        self.EmailPassword = os.getenv("SenderPassword")
        self.EFTUserName = os.getenv("EFTUser")
        self.EFTPassword = os.getenv("EFTPwd")

        # 呼叫Config
        self.GlobalConfig = Config()
        self.DealerConfig = DealerConf()
        self.MappingConfig = MappingRule()
        self.CheckConfig = CheckRule()
        self.DealerFormatConfig = DealerFormatConf()
        self.MailConfig = MailRule()
        self.UserConfig = User()
        self.FileConfig = File()

        # 測試模式
        self.TestMode = self.GlobalConfig["Default"]["TestMode"]

        # 錯誤嘗試次數
        self.MaxTryRange = self.GlobalConfig["Default"]["MaxTryRange"]

        # 工作日參數
        self.WorkDay = self.GlobalConfig["Default"]["WorkDay"]

        # 月繳檔案繳交區間
        self.MonthlyFileDeadline = self.GlobalConfig["Default"]["MonthlyFileDeadline"]

        # C槽
        self.SystemRoot = self.GlobalConfig["SystemRoot"]

        # OneDrive檔案名稱
        self.OneDeiveFolder = self.GlobalConfig["OneDirveFolder"]

        # OneDrive不拷貝目錄(包含資料夾都不拷貝)
        self.NotCopyFolder = self.GlobalConfig["NotCopyFolderFromOneDrive"]

        # 許可的副檔名
        self.AllowFileExtensions = self.GlobalConfig["Default"]["AllowFileExtensions"]

        # 取得Windows使用者名稱
        self.WinUser = self.GlobalConfig["App"]["WinUser"]

        # Log參數
        self.Operator = self.GlobalConfig["App"]["User"] \
                    if self.GlobalConfig["App"]["User"] \
                    else self.GlobalConfig["Default"]["User"]
        self.SystemLogFileName = self.GlobalConfig["LogConfig"]["SystemLog"]
        self.RecordLogFileName = self.GlobalConfig["LogConfig"]["RecordLog"]
        self.ChangeLogFileName = self.GlobalConfig["LogConfig"]["ChangeLog"]
        self.CheckLogFileName = self.GlobalConfig["LogConfig"]["CheckLog"]

        # Dealer清單
        self.DealerList = self.DealerConfig["DealerList"]
        self.KADealerList = self.DealerConfig["KADealerList"]

        # SMTP參數
        self.SMTPHost = self.GlobalConfig["Mail"]["SMTPHost"]
        self.SMTPPort = self.GlobalConfig["Mail"]["SMTPPort"]

        # 取得EFT Config
        self.EFTHostName = self.GlobalConfig["EFT"]["HostName"]
        self.EFTDir = self.GlobalConfig["EFT"]["Dir"]

        # MasterFile 檔案模板設定
        self.MasterFileName= self.GlobalConfig["MasterFile"]["FileName"]
        self.MasterFileSheetName = self.GlobalConfig["MasterFile"]["SheetName"]
        self.MasterFileHeader = self.GlobalConfig["MasterFile"]["Header"]
        self.MasterFileColumnWidth = self.GlobalConfig["MasterFile"]["ColumnWidth"]

        # KAList 檔案模板設定
        self.KAListFileName = self.GlobalConfig["KAList"]["FileName"]
        self.KAListFileSheetName = self.GlobalConfig["KAList"]["SheetName"]
        self.KAListFileHeader = self.GlobalConfig["KAList"]["Header"]
        self.KAListFileColumnWidth = self.GlobalConfig["KAList"]["ColumnWidth"]

        # DealerInfo 檔案模板設定
        self.DealerInfoFileName = self.GlobalConfig["DealerInfo"]["FileName"]
        self.DealerInfoFileSheetName = self.GlobalConfig["DealerInfo"]["SheetName"]
        self.DealerInfoFileHeader = self.GlobalConfig["DealerInfo"]["Header"]
        self.DealerInfoFileColumnWidth = self.GlobalConfig["DealerInfo"]["ColumnWidth"]


        # 檔案轉換規則
        self.SaleFileChangeRule = self.MappingConfig["MappingRule"]["Sale"]
        self.InventoryFileChangeRule = self.MappingConfig["MappingRule"]["Inventory"]

        # 銷售輸出參數
        self.SaleOutputFileName = self.GlobalConfig["OutputFile"]["Sale"]["FileName"]
        self.SaleOutputFileHeader = self.GlobalConfig["OutputFile"]["Sale"]["Header"]
        self.SaleOutputFileExtension = self.GlobalConfig["OutputFile"]["Sale"]["Extension"]
        self.SaleErrorReportFileName = self.GlobalConfig["ErrorReport"]["Sale"]["FileName"]

        # 庫存輸出參數
        self.InventoryOutputFileName = self.GlobalConfig["OutputFile"]["Inventory"]["FileName"]
        self.InventoryOutputFileHeader = self.GlobalConfig["OutputFile"]["Inventory"]["Header"]
        self.InventoryOutputFileExtension = self.GlobalConfig["OutputFile"]["Inventory"]["Extension"]
        self.InventoryOutputFileCountryCode = self.GlobalConfig["OutputFile"]["Inventory"]["CountryCode"]
        self.InventoryErrorReportFileName = self.GlobalConfig["ErrorReport"]["Inventory"]["FileName"]

        # 繳交紀錄表參數
        self.SubRawDataFileName = self.GlobalConfig["SubRawData"]["FileName"]
        self.SubRawDataFileName = self.SubRawDataFileName.replace("{Year}", str(self.Year))
        self.SubRawDataSheetName = self.GlobalConfig["SubRawData"]["SheetName"]
        self.SubRawDataSheetName = self.SubRawDataSheetName.replace("{Month}", str(self.Month))
        self.SubRawDataHeader = self.GlobalConfig["SubRawData"]["Header"]
        self.SubRawDataColumnWidth = self.GlobalConfig["SubRawData"]["ColumnWidth"]

        # DailyReport參數
        self.DailyReportFileName = self.GlobalConfig["DailyReport"]["FileName"]
        self.DailyReportFileName = self.DailyReportFileName.replace("{Year}", str(self.Year))
        self.DailyReportSheetName = self.GlobalConfig["DailyReport"]["SheetName"]
        self.DailyReportSheetName = self.DailyReportSheetName.replace("{Month}", str(self.Month))
        self.DailyReportHeader = self.GlobalConfig["DailyReport"]["Header"]
        self.DailyReportColumnWidth = self.GlobalConfig["DailyReport"]["ColumnWidth"]
        self.DailyReportNewDataWidth = self.GlobalConfig["DailyReport"]["NewDataWidth"]

        # MonthlyReport參數
        self.MonthlyReportFileName = self.GlobalConfig["MonthlyReport"]["FileName"]
        self.MonthlyReportFileName = self.MonthlyReportFileName.replace("{Year}", str(self.Year))
        self.MonthlyReportSheetName = self.GlobalConfig["MonthlyReport"]["SheetName"]
        self.MonthlyReportSheetName = self.MonthlyReportSheetName.replace("{Month}", str(self.Month))
        self.MonthlyReportHeader = self.GlobalConfig["MonthlyReport"]["Header"]
        self.MonthlyReportColumnWidth = self.GlobalConfig["MonthlyReport"]["ColumnWidth"]

        # 待補繳紀錄表參數
        self.NotSubFileName = self.GlobalConfig["NotSubmission"]["FileName"]
        self.NotSubFileName = self.NotSubFileName.replace("{Year}", str(self.Year))
        self.NotSubSheetName = self.GlobalConfig["NotSubmission"]["SheetName"]
        self.NotSubSheetName = self.NotSubSheetName.replace("{Month}", str(self.Month))
        self.NotSubHeader = self.GlobalConfig["NotSubmission"]["Header"]
        self.NotSubColumnWidth = self.GlobalConfig["NotSubmission"]["ColumnWidth"]

        # 銷售檔案參數
        self.SF_MustHave = self.CheckConfig["SaleFile"]["MustHave"]
        self.SF_2Choose1 = self.CheckConfig["SaleFile"]["2Choose1"]
        self.SF_Default_Header = self.DealerFormatConfig["Defualt"]["SaleFileHeader"]

        # 庫存檔案參數
        self.IF_MustHave = self.CheckConfig["InventoryFile"]["MustHave"]
        self.IF_2Choose1 = self.CheckConfig["InventoryFile"]["2Choose1"]
        self.IF_Default_Header = self.DealerFormatConfig["Defualt"]["InventoryFileHeader"]

        # 取得目錄參數
        self.OneDrivePath = self.GlobalConfig["App"]["OneDrivePath"]\
                        if self.GlobalConfig["App"]["OneDrivePath"]\
                        else self.GlobalConfig["Default"]["OneDrivePath"]
        self.OneDrivePath = self.OneDrivePath.replace("{username}", self.WinUser)
        self.RootDir = self.GlobalConfig["DirTree"]["Path"]
        self.FolderName = self.GlobalConfig["App"]["Name"] \
                    if self.GlobalConfig["App"]["Name"] \
                    else self.GlobalConfig["Default"]["Name"]
        self.SystemFolder = self.GlobalConfig["DirTree"]["System"]["FolderName"]
        self.ConfigFolder = self.GlobalConfig["DirTree"]["System"]["NextFolder"]["ConfigFolder"]
        self.BDFolder = self.GlobalConfig["DirTree"]["BD"]["FolderName"]
        self.MasterFileFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["MasterFileFolder"]
        self.ReportFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["ReportFolder"]["FolderName"]
        self.ErrorReportFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["ReportFolder"]["NextFolder"]["ErrorReportFolder"]
        self.BAFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["BAFolder"]["FolderName"]
        self.DealerInfoFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["DealerFolder"]
        self.DealerFolder = self.GlobalConfig["DirTree"]["Dealer"]["FolderName"]
        self.ChangeFolder = self.GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["ChangeFileFolder"]["FolderName"]
        self.MergeInventoryFolder = self.GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["ChangeFileFolder"]["NextFolder"]["MergeInventoryFolder"]
        self.CompleteFolder = self.GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["DealerFile"]["NextFolder"]["CompletedFolder"]

        # 制定全域目錄參數
        self.ConfigFolderPath = os.path.join(self.RootDir, self.FolderName, self.SystemFolder, self.ConfigFolder)
        self.BDFolderPath = os.path.join(self.RootDir, self.FolderName, self.BDFolder)
        self.MasterFolderPath = os.path.join(self.BDFolderPath, self.MasterFileFolder)
        self.ReportFolderPath = os.path.join(self.BDFolderPath, self.ReportFolder)
        self.ErrorReportPath = os.path.join(self.ReportFolderPath, self.ErrorReportFolder)
        self.SubRawDataPath = os.path.join(self.ReportFolderPath, self.SubRawDataFileName)
        self.DailyReportPath = os.path.join(self.ReportFolderPath, self.DailyReportFileName)
        self.MonthlyReportPath = os.path.join(self.ReportFolderPath, self.MonthlyReportFileName)
        self.NotSubPath = os.path.join(self.ReportFolderPath, self.NotSubFileName)
        self.BAFolderPath = os.path.join(self.BDFolderPath, self.BAFolder)
        self.DealerInfoPath = os.path.join(self.BDFolderPath, self.DealerInfoFolder)
        self.DealerFolderPath = os.path.join(self.RootDir, self.FolderName, self.DealerFolder)
        self.ChangeFolderPath = os.path.join(self.DealerFolderPath, self.ChangeFolder)

        # 取得路徑
        self.LogPath = self.GlobalConfig["LogConfig"]["Path"]
        self.TemplateFolderPath = self.GlobalConfig["MailTemplate"]
        
