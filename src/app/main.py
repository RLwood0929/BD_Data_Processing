'''
檔案說明：主流程控制
Writer：Qian
'''

import os
from dotenv import load_dotenv

load_dotenv()

path = os.getenv("OneDrivePath")

print(path)