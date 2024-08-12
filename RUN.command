#!/bin/bash

# 將工作目錄更改為腳本所在目錄
cd "$(dirname "$0")"

# 檢查是否安裝了 Python 3
if ! command -v python3 &> /dev/null
then
    echo "錯誤：未找到 Python 3。請安裝 Python 3 後再試。"
    echo "您可以從 https://www.python.org/downloads/ 下載 Python 3。"
    exit 1
fi

# 檢查虛擬環境是否存在，如果不存在則創建
if [ ! -d "venv" ]; then
    echo "正在創建虛擬環境..."
    python3 -m venv venv
fi

# 激活虛擬環境
source venv/bin/activate

# 檢查 requirements.txt 是否存在
if [ ! -f "requirements.txt" ]; then
    echo "錯誤：未找到 requirements.txt 文件。"
    echo "請確保 requirements.txt 文件與此腳本在同一目錄。"
    deactivate
    exit 1
fi

# 安裝或更新依賴
echo "正在安裝/更新依賴..."
pip install -r requirements.txt

# 檢查主 Python 腳本是否存在
if [ ! -f "discogs_to_sheets.py" ]; then
    echo "錯誤：未找到 discogs_to_sheets.py 文件。"
    echo "請確保 discogs_to_sheets.py 文件與此腳本在同一目錄。"
    deactivate
    exit 1
fi

# 運行 Python 腳本
echo "正在啟動 Discogs to Sheets 程序..."
python3 discogs_to_sheets.py

# 停用虛擬環境
deactivate

echo "程序執行完畢。按任意鍵退出。"
read -n 1 -s