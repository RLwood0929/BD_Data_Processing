# -*- coding: utf-8 -*-

'''
檔案說明：撰寫系統log及檔案檢查紀錄
Writer：Qian
level 0 對應 debug      適用範圍:詳細的程序運行信息，通常用於問题診斷和調試
level 1 對應 info       適用範圍:程式運行紀錄，表明程式正常工作
level 2 對應 warning    適用範圍:表明可能出現的問題，但程式仍然正常運行
level 3 對應 error      適用範圍:由於嚴重的問題，部分功能未執行
level 4 對應 critical   適用範圍:嚴重錯誤，程式不能繼續運行
'''

import os
import logging
from SystemConfig import Config

GlobalConfig = Config()

# 系統log存放位置
LogPath = GlobalConfig["App"]["LogPath"] if GlobalConfig["App"]["LogPath"] \
    else GlobalConfig["Default"]["LogPath"]
Operator = GlobalConfig["App"]["User"] if GlobalConfig["App"]["User"] \
    else GlobalConfig["Default"]["User"]
SystemLogFileName = GlobalConfig["Default"]["SystemLogFileName"]
RecordLogFileName = GlobalConfig["Default"]["RecordLogFileName"]
ChangeLogFileName = GlobalConfig["Default"]["ChangeLogFileName"]
CheckLogFileName = GlobalConfig["Default"]["CheckLogFileName"]

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(event)s - %(message)s"
)

# 建立 SystemLog，紀錄系統日誌
SysPath = os.path.join(LogPath, SystemLogFileName)
SysFileHandler = logging.FileHandler(SysPath, encoding="UTF-8")
SysFormatter = logging.Formatter\
    ("%(asctime)s - %(levelname)s - %(event)s - %(operator)s - %(message)s")
SysFileHandler.setFormatter(SysFormatter)

SysLogger = logging.getLogger("App")
SysLogger.addHandler(SysFileHandler)

# 建立 RecordLog，紀錄檔案繳交日誌
RecPath = os.path.join(LogPath, RecordLogFileName)
RecFileHandler = logging.FileHandler(RecPath, encoding="UTF-8")
RecFormatter = logging.Formatter\
    ("%(asctime)s - %(levelname)s - %(event)s - %(DealerID)s - %(FileName)s - %(message)s")
RecFileHandler.setFormatter(RecFormatter)

RecLogger = logging.getLogger("Record")
RecLogger.addHandler(RecFileHandler)

# 建立 ChangeLog，紀錄轉換日誌
ChaPath = os.path.join(LogPath, ChangeLogFileName)
ChaFileHandler = logging.FileHandler(ChaPath, encoding="UTF-8")
ChaFormatter = logging.Formatter\
    ("%(asctime)s - %(levelname)s - %(event)s - \
%(DealerID)s - %(FileName)s - %(message)s")
ChaFileHandler.setFormatter(ChaFormatter)

ChaLogger = logging.getLogger("Change")
ChaLogger.addHandler(ChaFileHandler)

# 建立 CheckLog，紀錄檔案檢查日誌
ChePath = os.path.join(LogPath, CheckLogFileName)
CheFileHandler = logging.FileHandler(ChePath, encoding="UTF-8")
CheFormatter = logging.Formatter\
    ("%(asctime)s - %(levelname)s - %(event)s - \
%(DealerID)s - %(FileName)s - %(message)s")
CheFileHandler.setFormatter(CheFormatter)

CheLogger = logging.getLogger("Check")
CheLogger.addHandler(CheFileHandler)

# 撰寫 System Log 紀錄，Level：0、1、2、3、4
def WSysLog(Level, Event, Message):
    Level = str(Level)
    if Level == "0":
        SysLogger.debug("Debug Message.", extra = {"event":"log功能除錯", "operator":"System"})
        sys_message = "Writing debug log is finish."
    elif Level == "1":
        SysLogger.info(Message, extra = {"event":Event, "operator":Operator})
        sys_message = "Writing info log is finish."
    elif Level == "2":
        SysLogger.warning(Message, extra = {"event":Event, "operator":Operator})
        sys_message = "Writing warning log is finish."
    elif Level == "3":
        SysLogger.error(Message, extra = {"event":Event, "operator":Operator})
        sys_message = "Writing error log is finish."
    elif Level == "4":
        SysLogger.critical(Message, extra = {"event":Event, "operator":Operator})
        sys_message = "Writing critical log is finish."
    else:
        sys_message = "Level out of range."
    return sys_message

# 撰寫 Record Log 紀錄，Level：1、2
def WRecLog(Level, Event, DealerID, FileName, Message):
    Level = str(Level)
    if Level == "1":
        RecLogger.info(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        rec_message = "Writing info log is finish."
    elif Level == "2":
        RecLogger.warning(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        rec_message = "Writing warning log is finish."
    else:
        rec_message = "Level out of range."
    return rec_message

# 撰寫 Check Log 紀錄，Level：1、2
def WCheLog(Level, Event, DealerID, FileName, Message):
    Level = str(Level)
    if Level == "1":
        CheLogger.info(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        che_message = "Writing info log is finish."
    elif Level == "2":
        CheLogger.warning(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        che_message = "Writing warning log is finish."
    else:
        che_message = "Level out of range."
    return che_message

# 撰寫 Change Log 紀錄，Level：1、2
def WChaLog(Level, Event, DealerID, FileName, Message):
    Level = str(Level)
    if Level == "1":
        ChaLogger.info(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        cha_message = "Writing info log is finish."
    elif Level == "2":
        ChaLogger.warning(Message, extra = {"event":Event, "DealerID":DealerID, "FileName":FileName})
        cha_message = "Writing warning log is finish."
    else:
        cha_message = "Level out of range."
    return cha_message

if __name__ == "__main__":
    level = "0"
    event = ""
    message = ""
    WSysLog(level, event, message)