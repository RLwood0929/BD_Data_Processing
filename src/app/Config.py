import os
from datetime import datetime
from dotenv import load_dotenv
from openpyxl.styles import Alignment
from SystemConfig import Config, DealerConf, CheckRule, DealerFormatConf,\
                         MappingRule, MailRule, User

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

        # 測試模式
        self.TestMode = self.GlobalConfig["Default"]["TestMode"]

        # 錯誤嘗試次數
        self.MaxTryRange = self.GlobalConfig["Default"]["MaxTryRange"]

        # 月繳檔案繳交區間
        self.MonthlyFileRange = self.GlobalConfig["App"]["MonthlyFileRange"] \
                            if self.GlobalConfig["App"]["MonthlyFileRange"]\
                            else self.GlobalConfig["Default"]["MonthlyFileRange"]

        # 許可的副檔名
        self.AllowFileExtensions = self.GlobalConfig["Default"]["AllowFileExtensions"]

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

        # 檔案轉換規則
        self.SaleFileChangeRule = self.MappingConfig["MappingRule"]["Sale"]
        self.InventoryFileChangeRule = self.MappingConfig["MappingRule"]["Inventory"]

        # 銷售輸出參數
        self.SaleOutputFileName = self.GlobalConfig["OutputFile"]["Sale"]["FileName"]
        self.SaleOutputFileHeader = self.GlobalConfig["OutputFile"]["Sale"]["Header"]
        self.SaleOutputFileExtension = self.GlobalConfig["OutputFile"]["Sale"]["Extension"]
        self.SaleErrorReportFileName = self.GlobalConfig["ErrorReport"]["Sale"]["FileName"]
        # SaleErrorReportHeader = GlobalConfig["ErrorReport"]["Sale"]["Header"]

        # 庫存輸出參數
        self.InventoryOutputFileName = self.GlobalConfig["OutputFile"]["Inventory"]["FileName"]
        self.InventoryOutputFileHeader = self.GlobalConfig["OutputFile"]["Inventory"]["Header"]
        self.InventoryOutputFileExtension = self.GlobalConfig["OutputFile"]["Inventory"]["Extension"]
        self.InventoryOutputFileCountryCode = self.GlobalConfig["OutputFile"]["Inventory"]["CountryCode"]
        self.InventoryErrorReportFileName = self.GlobalConfig["ErrorReport"]["Inventory"]["FileName"]
        # InventoryErrorReportHeader = GlobalConfig["ErrorReport"]["Inventory"]["Header"]

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
        self.RootDir = self.GlobalConfig["DirTree"]["Path"]
        self.FolderName = self.GlobalConfig["App"]["Name"] \
                    if self.GlobalConfig["App"]["Name"] \
                    else self.GlobalConfig["Default"]["Name"]

        self.BDFolder = self.GlobalConfig["DirTree"]["BD"]["FolderName"]
        self.MasterFileFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["MasterFileFolder"]
        self.ReportFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["ReportFolder"]["FolderName"]
        self.ErrorReportFolder = self.GlobalConfig["DirTree"]["BD"]["NextFolder"]["ReportFolder"]["NextFolder"]["ErrorReportFolder"]

        self.DealerFolder = self.GlobalConfig["DirTree"]["Dealer"]["FolderName"]
        self.ChangeFolder = self.GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["ChangeFileFolder"]["FolderName"]
        self.CompleteFolder = self.GlobalConfig["DirTree"]["Dealer"]["NextFolder"]["DealerFile"]["NextFolder"]["CompletedFolder"]

        # 制定全域目錄參數
        self.BDFolderPath = os.path.join(self.RootDir, self.FolderName, self.BDFolder)
        self.MasterFolderPath = os.path.join(self.BDFolderPath, self.MasterFileFolder)
        self.ReportFolderPath = os.path.join(self.BDFolderPath, self.ReportFolder)
        self.ErrorReportPath = os.path.join(self.ReportFolderPath, self.ErrorReportFolder)
        self.SubRawDataPath = os.path.join(self.ReportFolderPath, self.SubRawDataFileName)
        self.DailyReportPath = os.path.join(self.ReportFolderPath, self.DailyReportFileName)
        self.MonthlyReportPath = os.path.join(self.ReportFolderPath, self.MonthlyReportFileName)
        self.NotSubPath = os.path.join(self.ReportFolderPath, self.NotSubFileName)
        self.DealerFolderPath = os.path.join(self.RootDir, self.FolderName, self.DealerFolder)
        self.ChangeFolderPath = os.path.join(self.DealerFolderPath, self.ChangeFolder)

        # 取得路徑
        self.LogPath = self.GlobalConfig["LogConfig"]["Path"]
        self.TemplateFolderPath = self.GlobalConfig["MailTemplate"]

        # 取得MasterFile檔案
        self.MasterFile = [file for file in os.listdir(self.MasterFolderPath) \
                    if os.path.isfile(os.path.join(self.MasterFolderPath, file))] 
