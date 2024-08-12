@echo off

REM 檢查虛擬環境是否存在，如果不存在則創建
if not exist venv (
    python -m venv venv
)

REM 激活虛擬環境
call venv\Scripts\activate

REM 安裝或更新依賴
pip install -r requirements.txt

REM 運行Python腳本
python discogs_to_sheets.py

REM 停用虛擬環境
deactivate
