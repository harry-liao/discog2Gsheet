import discogs_client
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import *
import re

# Discogs API配置
DISCOGS_USER_TOKEN = 'VYSpCsnQtNrDsWETLCxZSSMRjNcdRjPTNvTUpXpw'

# Google Sheets API配置
GOOGLE_SHEETS_CREDENTIALS_FILE = 'elated-guild-424207-g7-58c6c832a459.json'
GOOGLE_SHEET_NAME = '奮死庫存2024 的副本'

# Discogs API客户端
d = discogs_client.Client('ExampleApplication/0.1', user_token=DISCOGS_USER_TOKEN)

def get_release_format(release):
    formats = release.formats
    if formats:
        return formats[0]['name']  # 返回第一个格式名称
    return 'Unknown'

def search_and_process(barcode):
    try:
        search_results = d.search(barcode=barcode, type='release')
        
        if search_results.count == 0:
            print(f'未找到條形碼對應的專輯: {barcode}')
            return None
        
        release = search_results[0]
        
        artists = ', '.join(artist.name for artist in release.artists)
        
        album_data = [
            1,  # 庫存，新數據設為1
            '',  # 價格，留空
            artists[0].upper() if artists else '',  # 開頭，取 Artists 第一個字母
            artists,
            release.title,
            '',  # 安娜murmur，留空
            release.year,
            ', '.join(label.name for label in release.labels),
            '',  # 推薦，留空
            get_release_format(release),
            ', '.join(release.genres),
            ', '.join(release.styles)
        ]
        
        return album_data
    except Exception as e:
        print(f'錯誤：處理條形碼 "{barcode}" 時發生問題，請嘗試其他格式，錯誤信息：{str(e)}')
        return None

def get_insertion_index(sheet, new_data):
    # 獲取C列（索引2）的所有值
    c_column = sheet.col_values(3)  # 3 對應 C 列
    
    # 如果新數據的C列為空，插入到最後
    if not new_data[2]:
        return len(c_column) + 1
    
    # 二分搜索找到插入位置
    left, right = 1, len(c_column)  # 從1開始，跳過標題行
    while left < right:
        mid = (left + right) // 2
        if c_column[mid] == '':  # 空值應該排在最後
            right = mid
        elif c_column[mid].lower() <= new_data[2].lower():
            left = mid + 1
        else:
            right = mid
    
    return left + 1  # +1 因為行號從1開始，而不是0

def update_or_add_row(sheet, new_data):
    # 檢查是否已存在相同的專輯
    existing_cells = sheet.findall(new_data[4], in_column=5)  # 在 E 列（專輯名稱）中查找
    
    if existing_cells:
        # 如果找到相同的專輯，更新庫存數量
        row = existing_cells[0].row
        current_stock = int(sheet.cell(row, 1).value)
        sheet.update_cell(row, 1, current_stock + 1)
        print(f'Updated stock for existing album: {new_data[4]}')
    else:
        # 如果是新專輯，找到正確的插入位置並添加
        insert_index = get_insertion_index(sheet, new_data)
        sheet.insert_row(new_data, insert_index)
        print(f'Added new album: {new_data[4]} at row {insert_index} ')
    
    # 只格式化新插入或更新的行
    format_new_row(sheet, insert_index if not existing_cells else row)

def format_new_row(sheet, row_index):
    # 設置條件格式
    colors = [
        {"red": 252/255, "green": 228/255, "blue": 205/255},  # 淺橘色
        {"red": 201/255, "green": 218/255, "blue": 248/255}   # 淺藍色
    ]
    
    # 獲取C列中直到新行的唯一值數量
    c_values = sheet.col_values(3)[:row_index]
    unique_count = len(set(c_values[1:]))  # 跳過標題行
    
    # 決定顏色
    color_index = unique_count % len(colors)
    color = colors[color_index]
    
    # 應用顏色到新行
    format_cell_range(sheet, f'A{row_index}:L{row_index}', CellFormat(backgroundColor=color))
    
    # 應用邊框
    format_cell_range(sheet, f'A{row_index}:L{row_index}', CellFormat(
        borders=Borders(
            top=Border(style='SOLID'),
            bottom=Border(style='SOLID'),
            left=Border(style='SOLID'),
            right=Border(style='SOLID')
        )
    ))

def main():
    # 認證 Google Sheets API
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(GOOGLE_SHEETS_CREDENTIALS_FILE, scopes=scope)
    client = gspread.authorize(creds)

    # 打開 Google Sheet
    sheet = client.open(GOOGLE_SHEET_NAME).worksheet('工作表1')

    while True:
        barcode = input("請輸入條碼 (輸入 'q' 退出): ").strip()
        if barcode.lower() == 'q':
            break
        
        album_data = search_and_process(barcode)
        
        if album_data:
            update_or_add_row(sheet, album_data)
            print('Data processed and sheet updated successfully!')
        else:
            print('No data to process.')

if __name__ == "__main__":
    main()