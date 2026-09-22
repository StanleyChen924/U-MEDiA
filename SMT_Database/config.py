import configparser
import os
import re
import shutil
#import sndhdr
import threading
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext  # 引入滾動文字方塊元件
from datetime import datetime
import glob
import pandas as pd


# ==============================================================================
# 0. 環境路徑與設定檔定義
# ==============================================================================
INI_FILENAME = 'MySQLConfig.ini'
SCAN_DIR = './ScanFolder'     # 要掃描的目錄
BACKUP_DIR = './BackupFolder' # 讀取成功後的備份目錄
LOG_FILENAME = f"./LOG/system_log{datetime.now().strftime('%Y%m%d')}.txt" # 記錄日誌的檔案名稱

# 1. 宣告設定檔讀取器
config = configparser.ConfigParser()
config_file_path = 'config.ini'

# 2. 檢查 config.ini 是否存在，存在就讀取，不存在就自動生成預設值（防呆）
if os.path.exists(config_file_path):
    config.read(config_file_path, encoding='utf-8')
else:
    # 如果現場人員不小心把 config.ini 刪除了，自動建立一個預設的
    config['PATH'] = {
        'SCAN_DIR': './ScanFolder',
        'BACKUP_DIR': './BackupFolder',
        'LOG_FILENAME': LOG_FILENAME
    }
    with open(config_file_path, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

# 3. 從 config.ini 中讀取路徑
# 使用 .get('區段名稱', '設定鍵名稱')，並加上 .strip() 移除多餘空格
SCAN_DIR = config.get('PATH', 'SCAN_DIR').strip()
BACKUP_DIR = config.get('PATH', 'BACKUP_DIR').strip()

# 4. 自動防呆：如果資料夾不存在，就自動建立它，避免程式崩潰
if not os.path.exists(SCAN_DIR):
    os.makedirs(SCAN_DIR)
    print(f"📁 偵測到掃描目錄不存在，已自動建立: {SCAN_DIR}")

if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)

os.makedirs(os.path.dirname(LOG_FILENAME), exist_ok=True)


# 自動建立掃描與備份所需的資料夾
os.makedirs(SCAN_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

# 全局變數：用於在背景執行緒與 UI 文字方塊之間傳遞訊息
monitor_text_area = None
