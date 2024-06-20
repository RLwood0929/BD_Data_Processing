# 開發文件

# 專案名稱：BD_Data_Processing

此為BD資料清洗專案主程式

## 開發工具及主要套件版本：

- Python：`3.11.9`

## 其他相關套件版本：

- python-dotenv：`1.0.1`

## 安裝套件：

```bash
pip install -r requirements.txt
```

## 建立虛擬環境：

```bash
#建置conda虛擬環境至系統預設目錄
conda create -n BD_DP python=3.11

#建置conda虛擬環境至指定目錄
conda create -p C:\coding\python\BD\DataProcessing\venv\BD_DP python=3.11
```

## 進入虛擬環境：

```bash
#建置conda虛擬環境至系統預設目錄
conda activate BD_DP

#建置conda虛擬環境至指定目錄
conda activate C:\coding\python\BD\DataProcessing\venv\BD_DP
```

## 資料夾結構

```markdown
DataProcessing/
├── datas/                      # 資料檔案存放位置
├── docs/                       # 開發的相關文件
|   ├── api/                    
|   ├── design/                 
|   └── user_guide/             
├── logs/                       # 系統運作log存放位置
├── src/                        
|   ├── app/                    
|   |   ├── __init.py__         
|   |   ├── check_file.py       # 紀錄檔案繳交時間，確認檔案格式
|   |   ├── EFT_file.py         # 檔案上傳至EFT雲端
|   |   ├── log.py              # 撰寫系統log及檔案檢查紀錄
|   |   ├── mail.py             # 寄送Mail
|   |   ├── main.py             # 主程式
|   |   ├── OneDrive_file.py    # OneDrive雲端檔案下載及上傳
|   |   └── window.py           # 視窗相關元件
|   ├── config/                 
|   ├── static                  
|   ├── templates/              
|   └── utils/                  
├── test/                       # 存放程式測試檔案
|   ├── integration             
|   └── unit                    
├── venv/                       # 存放虛擬環境
├── .env                        
├── .gitgnore                   
└──README.md                    
```

## 雲端OenDrive結構

```markdown
BD_DataProcessing/
├── 00_System/                  # 系統目錄
|   ├── Config/                 # 設定資料檔
|   |   ├── User_Config.csv     # 使用者config
|   |   ├── BD_Config.csv       # BD config
|   |   └── Dealer_Config.csv   # 經銷商 config
|   └─ Log/                     # 存放日誌記錄
|       ├── Success_Log         # 成功的日誌
|       └── Error_Log           # 失敗的日誌
├── 01_BD/                      # BD目錄
|   ├── 01_MainFile/            # 主檔存放位置
|   └── 02_Report/              # 報表存放位置
└── 02_Dealer/                  # 經銷商目錄
    ├── 00_CheckResult/         # 檢查結果存放位置
    ├── Dealer-1
    ├── Dealer-2
    ├── Dealer-3
    ├── Dealer-4
    ├── Dealer-5
    ├── Dealer-6
    ├── Dealer-7   
    └── Dealer-8
```

自定義函數名稱規則:

若此函數僅在單py檔中被呼叫則命名為小寫

若此函數會於其他py檔中被呼叫，則命名為大駝峰式