from DealerFileChange import change_file_name, split_file_data
from CheckFile import CheckFileHeader, CheckFileContent

dealer_id = "1002322861"
folder_path = "./datas"
file_key_word = "NCOSCA"
file_names = [
    "1002322861_S_202408310838.csv",
    "1002322861_S_202409300839.csv"
]

def DataProcessed():
    # 檔案名稱變更
    change_file_name(dealer_id, folder_path, file_key_word)

    # 檔案內容依據 Transion Date 篩選
    split_file_data(folder_path)

def CheckFileData():

    for file_name in file_names:

        print(f"file_name:{file_name}")

        # 檢查檔案表頭
        header_result = CheckFileHeader(dealer_id, file_name, "S")
        print(f"header_result:{header_result}")

        if header_result:
            # 檢查檔案內容
            content_result, error_num = CheckFileContent(dealer_id, file_name, "S")

            print(f"content_result:{content_result}")
            print(f"error_num:{error_num}")


# 線下轉換主流程
def OffLineWorkFlow():
    # DataProcessed()
    CheckFileData()


if __name__ == "__main__":
    OffLineWorkFlow()
