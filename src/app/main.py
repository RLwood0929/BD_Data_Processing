import os
from dotenv import load_dotenv

load_dotenv()

path = os.getenv("OneDrivePath")

print(path)