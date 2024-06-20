# -*- coding: utf-8 -*-

'''
檔案說明：讀取、變更System.json檔案
Writer：Qian
'''

import os
import json

ConfigDir = "src/config"
SystemConfig = "system.json"
MappingConfig = "mapping_rule.json"

ConfigPath = os.path.join(ConfigDir, SystemConfig)
MappingPath = os.path.join(ConfigDir, MappingConfig)

def Config():
    with open(ConfigPath,"r",encoding="utf-8") as file:
        config = json.load(file)
    return config

def MappingRule():
    with open(MappingPath, "r", encoding = "UTF-8")as file:
        MappingConfig = json.load(file)
    return MappingConfig