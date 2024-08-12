# -*- coding: utf-8 -*-

"""
檔案說明：Config檔案參數
Writer：Qian
"""

# 標準庫
import os

class ConfigJsonFile:
    CONFIG_FOLDER = "./src/config"
    
    # 定義配置文件名作為類屬性
    SYSTEM_CONFIG = "system.json"
    # 轉換規範
    MAPPING_CONFIG = "mapping.json"
    # 檔案檢查規範
    CHECK_RULE_CONFIG = "check_rule.json"
    # 經銷商資訊
    DEALER_CONFIG = "dealer.json"
    # 經銷商檔案表頭規定
    DEALER_FORMAT_CONFIG = "dealer_format.json"
    # 經銷商表頭傳換對照
    HEADER_CHANGE_CONFIG = "header_change.json"
    # 經銷商檔案繳交紀錄
    SUB_RECORD_CONFIG = "sub_record.json"
    # 信件寄件設定
    MAIL_RULE_CONFIG = "mail.json"
    # 使用者記錄檔
    USER_CONFIG = "user.json"
    # 檔案異動時間紀錄
    FILE_CONFIG = "files.json"

    def __init__(self):
        self.ConfigPath = os.path.join(self.CONFIG_FOLDER, self.SYSTEM_CONFIG)
        self.MappingPath = os.path.join(self.CONFIG_FOLDER, self.MAPPING_CONFIG)
        self.CheckPath = os.path.join(self.CONFIG_FOLDER, self.CHECK_RULE_CONFIG)
        self.DealerPath = os.path.join(self.CONFIG_FOLDER, self.DEALER_CONFIG)
        self.DealerFormatPath = os.path.join(self.CONFIG_FOLDER, self.DEALER_FORMAT_CONFIG)
        self.HeaderChangePath = os.path.join(self.CONFIG_FOLDER, self.HEADER_CHANGE_CONFIG)
        self.SubRecordPath = os.path.join(self.CONFIG_FOLDER, self.SUB_RECORD_CONFIG)
        self.MailRulePath = os.path.join(self.CONFIG_FOLDER, self.MAIL_RULE_CONFIG)
        self.UserConfigPath = os.path.join(self.CONFIG_FOLDER, self.USER_CONFIG)
        self.FileConfigPath = os.path.join(self.CONFIG_FOLDER, self.FILE_CONFIG)
