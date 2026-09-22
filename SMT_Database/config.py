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

# 防呆：如果 INI 檔案不存在，自動建立一份含初始值的設定檔
if not os.path.exists(INI_FILENAME):
    ini_content = """[SystemID]
Device_ID=0x0013
Vendor_ID=0x168C
SSYS_ID=0x2051
SSYS_VEND_ID=0x168C

[setting]
MySQL_FLAG=1
MySQL_InsertFlag=1
MySQL_BeforStation=and STA1 > 100
TableDetailStr=STA1
MySQL_ServerIP=10.4.5.13
MySQL_username=U94003
MySQL_Password=U94003
MySQL_DB=LITEON
MySQL_TYPE=CARD
MySQL_Job=PUA00K5
MySQL_ModelName=BIW-LA00
MySQL_Operator=U93039
MySQL_Station=1006-1001
MySQL_CPNumber=4-3-1
"""
    with open(INI_FILENAME, 'w', encoding='utf-8') as f:
        f.write(ini_content)

# ==============================================================================
# 1. 核心邏輯：讀取與寫入 INI 檔案 / 日誌記錄
# ==============================================================================
def load_config():
    """精準載入 INI，保持原始大小寫"""
    config = configparser.ConfigParser(
        comment_prefixes=('#', ';', '//'),
        inline_comment_prefixes=('#', ';', '//'),
        strict=False
    )
    config.optionxform = str
    config.read(INI_FILENAME, encoding='utf-8')
    return config


def save_config(config, job, model, operator, station):
    """將修改後的 4 個數值覆寫回原 INI，並完美保留所有 `//` 註解"""
    try:
        with open(INI_FILENAME, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        targets = {
            'MySQL_Job': job, 'MySQL_ModelName': model,
            'MySQL_Operator': operator, 'MySQL_Station': station
        }

        new_lines = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith('//') or stripped.startswith('#') or stripped.startswith(';'):
                new_lines.append(line)
                continue

            replaced = False
            for key, val in targets.items():
                # 使用分割符來精準匹配 key，避免空格影響
                if '=' in stripped:
                    current_key = stripped.split('=', 1)[0].strip()
                    if current_key == key:
                        new_lines.append(f"{key}={val}\n")
                        replaced = True
                        break
            if not replaced:
                new_lines.append(line)

        with open(INI_FILENAME, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        return True
    except Exception as e:
        messagebox.showerror("錯誤", f"存檔失敗：{e}")
        return False


def mysql_safe_identifier(name, field_name="identifier"):
    """驗證 SQL 識別字（表名/欄位名）避免 `%s` 風險與注入問題。"""
    if name is None:
        raise ValueError(f"{field_name} 不可為空")

    value = str(name).strip()
    if not re.fullmatch(r"[A-Za-z0-9_]+", value):
        raise ValueError(f"{field_name} 含有非法字元：{name}")
    return value


def log_and_display(message):
    """將訊息寫入 LOG 檔，並即時顯示在監控視窗的 Text 欄位中"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_message = f"[{timestamp}] {message}"

    # 1. 寫入本地 LOG 檔案
    try:
        os.makedirs(os.path.dirname(LOG_FILENAME), exist_ok=True)
        with open(LOG_FILENAME, 'a', encoding='utf-8') as log_file:
            log_file.write(full_message + "\n")
    except Exception as e:
        print(f"寫入日誌檔失敗: {e}")

    # 2. 更新到 UI 的 Text 欄位
    if monitor_text_area is not None:
        try:
            monitor_text_area.insert(tk.END, full_message + "\n")
            monitor_text_area.see(tk.END) # 自動滾動到最底端
        except Exception:
            pass


# ==============================================================================
# 2. 讀取檔案：讀取所有excel檔案
# ==============================================================================
def normalize_column_name(column_name):
    """將 Excel 欄位名稱標準化，降低命名差異造成的比對失敗。"""
    if column_name is None:
        return ""

    text = str(column_name).strip().lower()
    text = text.replace(" ", "")
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", text)
    return text


def align_dataframe_columns(df):
    """將常見欄位別名自動對齊到標準欄位名稱。"""
    if df is None:
        return df

    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    alias_map = {
        "工单号": "工單號",
        "工單號": "工單號",
        "工单": "工單號",
        "工號": "工單號",
        "workorder": "工單號",
        "orderno": "工單號",
        "job": "工單號",
        "panelno": "PANEL_NO",
        "panel": "PANEL_NO",
        "panelnumber": "PANEL_NO",
        "panelnum": "PANEL_NO",
        "panel_no": "PANEL_NO",
        "pnl": "PANEL_NO",
        "併板主pnl": "併板主PNL",
        "主pnl": "併板主PNL",
        "pnlmain": "併板主PNL",
        "序號": "序號(SN)",
        "序號sn": "序號(SN)",
        "sn": "序號(SN)",
        "s/n": "序號(SN)",
        "serialnumber": "序號(SN)",
        "序號s1sn": "序號(S1SN)",
        "s1sn": "序號(S1SN)",
        "序號s2sn": "序號(S2SN)",
        "s2sn": "序號(S2SN)",
    }
    normalized_alias_map = {normalize_column_name(key): value for key, value in alias_map.items()}

    rename_map = {}
    for column in df.columns:
        normalized = normalize_column_name(column)
        target_name = normalized_alias_map.get(normalized)
        if not target_name:
            continue

        if target_name in df.columns and target_name != column:
            # 如果目標欄位已存在，優先保留原有欄位，將當前欄位值填入空白的地方
            df[target_name] = df[target_name].combine_first(df[column])
            continue

        rename_map[column] = target_name

    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def load_all_excel_files():
    # 1. 設定放置所有 .xlsx 檔案的資料夾路徑 ('.' 代表目前程式所在的資料夾)
    folder_path = '.'
    file_pattern = os.path.join(folder_path, '*.xlsx')

    # 2. 找出所有符合的 .xlsx 檔案
    excel_files = glob.glob(file_pattern)

    # 定義最終保留的標準欄位：保留原本欄位，並追加新增的判斷欄位
    required_columns = ['工單號']
    output_columns = ['工單號', 'PANEL_NO', '序號(SN)', '併板主PNL', '序號(S1SN)', '序號(S2SN)']

    # 3. 建立一個空的串列，用來存放每個檔案處理後的資料
    df_list = []

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        try:
            df = pd.read_excel(file_path, dtype={'PANEL_NO': str})

            # 清除欄位名稱的前後空白，避免因為「工單號 」有空格而對不上
            df.columns = df.columns.str.strip()

            # 自動對齊常見欄位別名，提升 Excel 相容性
            df = align_dataframe_columns(df)

            # 補齊舊欄位與新判斷欄位的相互對應，兼容兩種 Excel 格式
            if 'PANEL_NO' not in df.columns and '併板主PNL' in df.columns:
                df['PANEL_NO'] = df['併板主PNL']
            if '併板主PNL' not in df.columns and 'PANEL_NO' in df.columns:
                df['併板主PNL'] = df['PANEL_NO']

            if '序號(SN)' not in df.columns:
                if '序號(S1SN)' in df.columns:
                    df['序號(SN)'] = df['序號(S1SN)']
                elif '序號(S2SN)' in df.columns:
                    df['序號(SN)'] = df['序號(S2SN)']
            if '序號(S1SN)' not in df.columns and '序號(SN)' in df.columns:
                df['序號(S1SN)'] = df['序號(SN)']
            if '序號(S2SN)' not in df.columns and '序號(SN)' in df.columns:
                df['序號(S2SN)'] = df['序號(SN)']

            # 檢查該 Excel 是否包含最少必要欄位（工單號）
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                print(f"⚠️ 檔案 [{filename}] 缺少必須欄位 {missing_cols}，將跳過此檔案。")
                continue

            df_filtered = df.reindex(columns=output_columns).copy()

            # 【重要】將所有資料轉為字串並去除前後空白，防止比對時因格式不符（如數字/文字混雜）而失敗
            for col in output_columns:
                df_filtered[col] = df_filtered[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

            # 選擇性：記錄來源檔案，方便未來異常追溯
            df_filtered['來源檔案'] = filename

            df_list.append(df_filtered)

        except Exception as e:
            print(f"❌ 讀取檔案失敗 [{filename}]: {e}")

    # 4. 合併所有資料，形成最終的標準化 BUFF
    if df_list:
        data_buffer = pd.concat(df_list, axis=0, ignore_index=True)
        return data_buffer
    else:
        data_buffer = pd.DataFrame(columns=output_columns)
        print("\n❌ 錯誤：沒有成功讀取到任何符合欄位條件的 Excel 檔案。")
        return pd.DataFrame()


# 模式 A：使用 DataFrame 直接篩選比對（適合查出該 PANEL_NO 的完整整列資料）
def find_sn_by_panel_df(target_panel, buffer, BSN):
    # 防呆：如果全域 BUFF 本身是空的，直接回傳空表
    if buffer is None or buffer.empty:
        return pd.DataFrame()

    # 1. 將檔名切出來的 PANEL_NO 轉字串、去空白、轉大寫
    target_clean = str(target_panel).strip().upper()

    # 2. 將 BUFF 裡的 PANEL_NO 欄位也全部轉字串、去空白、轉大寫，再進行比對
    # 這樣可以防止 Excel 內部的「純數字格式」或「空格」導致比對失敗
    result = buffer[buffer['PANEL_NO'].astype(str).str.strip().str.upper() == target_clean].copy()

    if result.empty:
        return pd.DataFrame()

    result = result.reset_index(drop=True)

    # 3. 針對同一個 PANEL_NO，將原始資料展開成「S1SN -> S2SN」的順序
    #    這樣第 3、4 筆就會對應到 S2SN 的資料，而不是混在前兩筆之中
    expanded_rows = []

    # 先放 S1SN
    for _, row in result.iterrows():
        s1sn = row.get('序號(S1SN)', row.get('序號(SN)', ''))
        if pd.notna(s1sn):
            s1_row = row.copy()
            s1_row['序號(SN)'] = str(s1sn).strip()
            expanded_rows.append(s1_row)

    # 再放 S2SN
    for _, row in result.iterrows():
        s2sn = row.get('序號(S2SN)')
        if pd.notna(s2sn):
            s2_row = row.copy()
            s2_row['序號(SN)'] = str(s2sn).strip()
            expanded_rows.append(s2_row)

    expanded_result = pd.DataFrame(expanded_rows)

    # 4. 將傳入的 BSN 轉為整數
    try:
        target_index = int(BSN)
    except (TypeError, ValueError):
        print(f"⚠️ 警告：BSN [{BSN}] 非法，無法比對。")
        return pd.DataFrame()

    # 5. 防呆：檢查比對到的總筆數，是否足夠抓取指定的第幾筆
    if len(expanded_result) >= target_index:
        # 因為 Python 索引從 0 開始，所以「第三筆」的索引是 2
        record = expanded_result.iloc[[target_index - 1]].copy()
    else:
        print(
            f"⚠️ 警告：PANEL_NO [{target_clean}] 僅找到 {len(expanded_result)} 筆資料，無法獲取第 {target_index} 筆！"
        )
        record = pd.DataFrame()

    return record


# ==============================================================================
# 3. 背景任務：每 30 秒自動掃描與上拋 MySQL
# ==============================================================================
def parse_content_file(file_path):
    sn = "UNKNOWN"
    dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if line.upper().startswith("SN="):
                    sn = line.split("=", 1)[1]
                elif line.upper().startswith("DATETIME="):
                    dt = line.split("=", 1)[1]
    except Exception as e:
        print(f"讀取檔案 {file_path} 失敗: {e}")
    return sn, dt


def upload_to_mysql(sn_param, file_datetime, Side, status, cfg, MySQL_Job, panel_no, file_name, ModelName):
    if cfg['setting'].getint('MySQL_FLAG', 0) != 1:
        log_and_display("MySQL_FLAG 未啟用，跳過上拋。")
        return False

    max_retries = 3

    try:
        MySQL_TYPE = mysql_safe_identifier(cfg['setting'].get('MySQL_TYPE', 'CARD'), 'MySQL_TYPE')
        base_table = mysql_safe_identifier(cfg['setting'].get('TableDetailStr', 'STA1'), 'TableDetailStr')
        TableDetailStr = f"{base_table}_B" if Side == "B" else f"{base_table}_T"
    except ValueError as e:
        log_and_display(f"MySQL 表名設定錯誤: {e}")
        return False

    for attempt in range(1, max_retries + 1):
        conn = None
        try:
            if attempt == 1:
                log_and_display(
                    f"[重送上傳][SN={sn_param}][PANEL={panel_no}][Job={MySQL_Job}][DB寫入][第{attempt}/{max_retries}次]"
                )
            else:
                log_and_display(
                    f"[重送上傳][SN={sn_param}][PANEL={panel_no}][Job={MySQL_Job}][DB寫入][重試第{attempt}/{max_retries}次]"
                )

            conn = pymysql.connect(
                host=cfg['setting'].get('MySQL_ServerIP'),
                user=cfg['setting'].get('MySQL_username'),
                password=cfg['setting'].get('MySQL_Password'),
                database=cfg['setting'].get('MySQL_DB'),
                charset='utf8',
                connect_timeout=5
            )

            MySQL_InsertFlag = int(cfg['setting'].get('MySQL_InsertFlag', 0))

            with conn.cursor() as cursor:
                full_string = ""

                if MySQL_InsertFlag == 1:
                    sql = (
                        f"INSERT INTO `{MySQL_TYPE}` "
                        "SET iSN=%s, SMT_PN=%s, PANEL_SN=%s, "
                        f"`{TableDetailStr}`=101 "
                        "ON DUPLICATE KEY UPDATE "
                        "SMT_PN=%s, PANEL_SN=%s"
                    )
                    cursor.execute(
                        sql,
                        (sn_param, MySQL_Job, panel_no, MySQL_Job, panel_no)
                    )
                else:
                    sql = (
                        f"UPDATE `{MySQL_TYPE}` SET "
                        "SMT_PN=%s, PANEL_SN=%s, "
                        f"`{TableDetailStr}`=101 {full_string} "
                        "WHERE iSN = %s"
                    )
                    cursor.execute(sql, (MySQL_Job, panel_no, sn_param))

                MySQL_ModelName = ModelName
                MySQL_Operator = cfg['setting'].get('MySQL_Operator')
                MySQL_Station = cfg['setting'].get('MySQL_Station')

                sql = (
                    f"INSERT INTO `{TableDetailStr}` SET "
                    "iSN=%s, errorCode=%s, JobNum=%s, ModelName=%s, "
                    "operator=%s, Station=%s, StartTime=%s, StopTime=%s, "
                    "logfilename=%s, log=%s"
                )
                cursor.execute(
                    sql,
                    (
                        sn_param,
                        status,
                        MySQL_Job,
                        MySQL_ModelName,
                        MySQL_Operator,
                        MySQL_Station,
                        file_datetime,
                        file_datetime,
                        file_name,
                        file_name,
                    )
                )

            conn.commit()

            # 二次確認：確認主表狀態與明細表紀錄是否已成功寫入
            with conn.cursor() as verify_cursor:
                verify_main_sql = (
                    f"SELECT `{TableDetailStr}` FROM `{MySQL_TYPE}` WHERE iSN = %s"
                )
                verify_cursor.execute(verify_main_sql, (sn_param,))
                main_result = verify_cursor.fetchone()
                main_upload_success = main_result is not None and str(main_result[0]) == "101"

                verify_detail_sql = (
                    f"SELECT 1 FROM `{TableDetailStr}` "
                    "WHERE iSN = %s AND JobNum = %s AND logfilename = %s LIMIT 1"
                )
                verify_cursor.execute(
                    verify_detail_sql,
                    (sn_param, MySQL_Job, file_name)
                )
                detail_upload_success = verify_cursor.fetchone() is not None

            if main_upload_success and detail_upload_success:
                log_and_display(
                    f"[二次確認成功][SN={sn_param}][PANEL={panel_no}]"
                    f"[主表狀態={TableDetailStr}=101][明細紀錄存在]"
                )
                return True

            if not main_upload_success:
                log_and_display(
                    f"[二次確認失敗][SN={sn_param}][PANEL={panel_no}]"
                    f"[主表未確認為 {TableDetailStr}=101]"
                )

            if not detail_upload_success:
                log_and_display(
                    f"[二次確認失敗][SN={sn_param}][PANEL={panel_no}]"
                    f"[明細表無對應紀錄][Job={MySQL_Job}][File={file_name}]"
                )

            return False

        except Exception as e:
            log_and_display(
                f"[重送上傳][SN={sn_param}][DB寫入][失敗][第{attempt}/{max_retries}次][錯誤={e}]"
            )
            if attempt < max_retries:
                time.sleep(2)
                continue

            log_and_display(
                f"[重送上傳][SN={sn_param}][DB寫入][失敗][重試上限達成][共{max_retries}次]"
            )
            log_and_display(f"MySQL 連線或寫入錯誤: {e}")
            return False
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass


def scan_folder_loop():
    # 1. 先讀取所有的 Excel 檔案建立 BUFF
    while True:
        try:
            current_cfg = load_config()
            files = [f for f in os.listdir(SCAN_DIR) if os.path.isfile(os.path.join(SCAN_DIR, f))]
            if files:
                log_and_display(f"偵測到 {len(files)} 個新檔案，開始比對並上拋...")
                for file_name in files:
                    file_path = os.path.join(SCAN_DIR, file_name)

                    # 使用獨立 try-except 隔離單一檔案錯誤，避免整個迴圈中斷
                    try:
                        # 1. 解析檔名（假設格式：機型_工單_PANELNO_狀態_面別_時間.xlsx，請依實際情況微調索引）
                        result = file_name.split('_')

                        # 安全防呆：確保檔名資訊完整，避免 IndexError
                        if len(result) < 5:
                            log_and_display(f"⚠️ 檔案 [{file_name}] 名稱格式不符，跳過處理。")
                            # 仍將錯誤檔名搬移，避免卡在掃描區
                            shutil.move(file_path, os.path.join(BACKUP_DIR, f"ERR_NAME_{file_name}"))
                            continue

                        Station = result[0]
                        panel_no = result[2].strip()  # 取出 PANEL_NO 用於比對
                        status = result[3]
                        BSN = result[4]
                        ModelName = result[5][:-1]
                        target_field = result[-2]
                        Side = target_field[-1]

                        # 2. 處理時間欄位防呆（假設時間在 result[5]，請根據您真實的檔名位置修改索引數字）
                        time_part = os.path.splitext(result[-1])[0].strip().upper()

                        if time_part == 'NONE' or time_part == '':
                            file_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            log_and_display(f"⚠️ 檔案 [{file_name}] 時間為 NONE，自動代入目前系統時間。")
                        else:
                            try:
                                dt_object = datetime.strptime(time_part, "%Y%m%d%H%M%S")
                                file_dt = dt_object.strftime("%Y-%m-%d %H:%M:%S")
                            except ValueError:
                                file_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                log_and_display(f"⚠️ 檔案 [{file_name}] 時間格式錯誤，自動代入目前系統時間。")

                        # 3. 透過 PANEL_NO 到 Excel 緩衝區 (data_buffer) 比對找出 序號(SN)
                        df_res = find_sn_by_panel_df(panel_no, data_buffer, BSN)

                        # 明確排除 FAIL 狀態，其餘狀態皆可上拋
                        if status.upper() == "FAIL":
                            log_and_display(f"ℹ️ 檔案 [{file_name}] 狀態為 FAIL，跳過上拋資料庫。")
                        else:
                            if df_res is not None and not df_res.empty:
                                log_and_display(f" PANEL_NO [{panel_no}] 比對成功，找到 {len(df_res)} 筆對應的 SN 資料，開始逐筆上拋...")

                                # 4. 走訪比對到的每一筆資料，將資料全部上拋到資料庫
                                for _, row in df_res.iterrows():
                                    actual_sn = str(row['序號(SN)']).strip()
                                    work_order = result[1]
                                    MySQL_Job = work_order
                                    success = upload_to_mysql(actual_sn, file_dt, Side, status, current_cfg, MySQL_Job, panel_no, file_name, ModelName)

                                    if success:
                                        log_and_display(f"  └─ 成功: PANEL [{panel_no}] -> 轉出 SN: {actual_sn} 上拋成功")
                                    else:
                                        log_and_display(f"  └─ 失敗: PANEL [{panel_no}] -> 轉出 SN: {actual_sn} 上拋失敗")
                            else:
                                log_and_display(f"⚠️ 比對失敗：在 Excel 緩衝區中找不到 PANEL_NO [{panel_no}] 的任何資料，此檔案不上拋。")

                    except Exception as file_e:
                        log_and_display(f"⚠️ 處理個別檔案 {file_name} 時發生異常: {file_e}")

                    # 5. 無論上拋成功與否（或沒比對到），皆搬移檔案至備份目錄，保持目錄清空
                    dest_path = os.path.join(BACKUP_DIR, file_name)

                    if os.path.exists(dest_path):
                        base, ext = os.path.splitext(file_name)
                        timestamp = datetime.now().strftime("%H%M%S")
                        dest_path = os.path.join(BACKUP_DIR, f"{base}_{timestamp}{ext}")

                    try:
                        shutil.move(file_path, dest_path)
                    except Exception as move_e:
                        log_and_display(f"搬移檔案 {file_name} 失敗: {move_e}")
            else:
                # 減少日誌噪音
                pass
        except Exception as e:
            log_and_display(f"⚠️ 掃描迴圈最外層發生嚴重錯誤: {e}")

        time.sleep(30)


# ==============================================================================
# 4. UI 介面控制邏輯
# ==============================================================================
def open_monitor_ui():
    """【第二階段】開啟自動掃描監控 UI 視窗 (含文字顯示與滾動條)"""
    global monitor_text_area

    monitor_root = tk.Tk()
    monitor_root.title("系統執行監控中")
    monitor_root.geometry("550x380")
    monitor_root.resizable(False, False)

    # 標題
    tk.Label(monitor_root, text="🚀 系統自動監控服務中", font=("微軟正黑體", 12, "bold"), fg="green").pack(pady=10)

    # ScrolledText 欄位：即時顯示掃描與檔案資訊
    monitor_text_area = scrolledtext.ScrolledText(monitor_root, width=65, height=14, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4")
    monitor_text_area.pack(padx=15, pady=5)

    log_and_display("監控服務啟動成功。")

    def on_exit():
        """點擊 EXIT 直接安全離開整個程式"""
        monitor_root.destroy()

    # EXIT 按鈕
    btn_exit = tk.Button(monitor_root, text="EXIT (停止並離開系統)", bg="#333333", fg="white", font=("微軟正黑體", 10, "bold"), width=22, command=on_exit)
    btn_exit.pack(pady=15)

    # 啟動 30 秒定時掃描執行緒
    bg_thread = threading.Thread(target=scan_folder_loop, daemon=True)
    bg_thread.start()

    monitor_root.mainloop()


def create_setup_ui():
    """【第一階段】設定 INI 的初始 UI 視窗"""
    global root, entries, config_data, excel_data, data_buffer

    config_data = load_config()
    setting = config_data['setting'] if 'setting' in config_data else {}

    root = tk.Tk()
    root.title("MySQL 參數設定調整")
    root.geometry("400x260")
    root.resizable(False, False)

    padx_val = 15
    pady_val = 6

    labels = ["MySQL_Job:", "MySQL_ModelName:", "MySQL_Operator:", "MySQL_Station:"]
    entries = {}
    keys = ["MySQL_Job", "MySQL_ModelName", "MySQL_Operator", "MySQL_Station"]

    for i, (label_text, key) in enumerate(zip(labels, keys)):
        tk.Label(root, text=label_text, font=("Arial", 10, "bold")).grid(row=i, column=0, sticky="e", padx=padx_val, pady=pady_val)
        entry = tk.Entry(root, width=30)
        entry.grid(row=i, column=1, padx=padx_val, pady=pady_val)
        entry.insert(0, setting.get(key, ''))
        entries[key] = entry

    def on_confirm():
        success = save_config(
            config_data,
            entries['MySQL_Job'].get().strip(),
            entries['MySQL_ModelName'].get().strip(),
            entries['MySQL_Operator'].get().strip(),
            entries['MySQL_Station'].get().strip()
        )
        if success:
            root.destroy()
            open_monitor_ui()

    def on_cancel():
        root.destroy()
        open_monitor_ui()

    btn_frame = tk.Frame(root)
    btn_frame.grid(row=4, column=0, columnspan=2, pady=15)

    btn_cancel = tk.Button(btn_frame, text=" 否 (不存檔直接監控) ", bg="#d9534f", fg="white", font=("微軟正黑體", 9, "bold"), width=16, command=on_cancel)
    btn_cancel.pack(side="left", padx=15)

    btn_confirm = tk.Button(btn_frame, text="確認 (存檔並監控)", bg="#5cb85c", fg="white", font=("微軟正黑體", 9, "bold"), width=16, command=on_confirm)
    btn_confirm.pack(side="right", padx=15)

    root.mainloop()


# ==============================================================================
# 主程式入口 (包含環境檢測與防閃退)
# ==============================================================================
if __name__ == "__main__":
    # 讀取 Excel 建立 BUFF
    data_buffer = load_all_excel_files()

    try:
        # 動態載入 pymysql
        import pymysql
    except ImportError:
        # 如果使用者環境中沒裝 pymysql，使用基礎視窗警告防閃退
        root = tk.Tk()
        root.title("警告")
        root.geometry("300x100")
        root.resizable(False, False)
        tk.Label(root, text="警告：您未安裝 pymysql 模組，無法繼續執行。", font=("Arial", 10, "bold"), fg="red").pack(pady=20)
        root.mainloop()
        exit(1)

    # 執行主 UI
    create_setup_ui()
