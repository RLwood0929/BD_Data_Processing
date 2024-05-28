# 開發文件

# 專案名稱：BD_Data_Processing

此為BD資料清洗專案主程式

## 開發工具及主要套件版本：

- Python：`3.11.9`

## 其他相關套件版本：

```python

```

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
├── datas/                  # 資料檔案存放位置
├── docs/                   # 開發的相關文件
|   ├── api/                
|   ├── design/             
|   └── user_guide/         
├── logs/                   # 系統運作log存放位置
├── src/                    
|   ├── app/                
|   |   ├── __init.py__      
|   |   └── main.py         # 主程式
|   ├── config/             
|   ├── static              
|   ├── templates/          
|   └── utils/              
├── test/                   # 存放程式測試檔案
|   ├── integration         
|   └── unit                
├── venv/                   # 存放虛擬環境
├── .env                    
├── .gitgnore               
└──README.md                
```