import os

class ConfigJsonFile:
    CONFIG_FOLDER = "./src/config"
    
    # 定義配置文件名作為類屬性
    SYSTEM_CONFIG = "system.json"
    MAPPING_CONFIG = "mapping.json"
    CHECK_RULE_CONFIG = "check_rule.json"
    DEALER_CONFIG = "dealer.json"
    DEALER_FORMAT_CONFIG = "dealer_format.json"
    HEADER_CHANGE_CONFIG = "header_change.json"
    SUB_RECORD_CONFIG = "sub_record.json"
    MAIL_RULE_CONFIG = "mail.json"
    USER_CONFIG = "user.json"
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
