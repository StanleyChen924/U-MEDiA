import configparser
import glob
import os
import re
import shutil
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, scrolledtext

import pandas as pd


# ==============================================================================
# 環境路徑與設定檔
# ==============================================================================
INI_FILENAME = 'MySQLConfig.ini'
config_file_path = 'config.ini'

config = configparser.ConfigParser()
if os.path.exists(config_file_path):
    config.read(config_file_path, encoding='utf-8')
else:
    config['PATH'] = {
        'SCAN_DIR': './ScanFolder',
        'BACKUP_DIR': './BackupFolder',
        'FAIL_DIR': './FailFolder',
        'LOG_FILENAME': f"./LOG/system_log{datetime.now().strftime('%Y%m%d')}.txt",
        'FAIL_LOG_FILENAME': f"./LOG/FIAL_system_log{datetime.now().strftime('%Y%m%d')}.txt",
    }
    with open(config_file_path, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

PATH_CONFIG = config['PATH']
SCAN_DIR = PATH_CONFIG.get('SCAN_DIR', './ScanFolder').strip()
BACKUP_DIR = PATH_CONFIG.get('BACKUP_DIR', './BackupFolder').strip()
FAIL_DIR = PATH_CONFIG.get('FAIL_DIR', './FailFolder').strip()
LOG_FILENAME = PATH_CONFIG.get(
    'LOG_FILENAME',
    f"./LOG/system_log{datetime.now().strftime('%Y%m%d')}.txt",
).strip()
FAIL_LOG_FILENAME = PATH_CONFIG.get(
    'FAIL_LOG_FILENAME',
    f"./LOG/FIAL_system_log{datetime.now().strftime('%Y%m%d')}.txt",
).strip()

for directory in (SCAN_DIR, BACKUP_DIR, FAIL_DIR, os.path.dirname(LOG_FILENAME), os.path.dirname(FAIL_LOG_FILENAME)):
    if directory:
        os.makedirs(directory, exist_ok=True)

monitor_text_area = None

if not os.path.exists(INI_FILENAME):
    with open(INI_FILENAME, 'w', encoding='utf-8') as f:
        f.write("""[SystemID]
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
""")


# ==============================================================================
# 設定、日誌與檔案搬移
# ==============================================================================
def load_config():
    result = configparser.ConfigParser(
        comment_prefixes=('#', ';', '//'),
        inline_comment_prefixes=('#', ';', '//'),
        strict=False,
    )
    result.optionxform = str
    result.read(INI_FILENAME, encoding='utf-8')
    return result


def save_config(config_data, job, model, operator, station):
    try:
        with open(INI_FILENAME, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        targets = {
            'MySQL_Job': job,
            'MySQL_ModelName': model,
            'MySQL_Operator': operator,
            'MySQL_Station': station,
        }
        new_lines = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith(('/', '#', ';')) or '=' not in stripped:
                new_lines.append(line)
                continue
            key = stripped.split('=', 1)[0].strip()
            new_lines.append(f"{key}={targets[key]}\n" if key in targets else line)
        with open(INI_FILENAME, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        return True
    except Exception as error:
        messagebox.showerror('錯誤', f'存檔失敗：{error}')
        return False


def mysql_safe_identifier(value, field_name='identifier'):
    value = str(value or '').strip()
    if not re.fullmatch(r'[A-Za-z0-9_]+', value):
        raise ValueError(f'{field_name} 含有非法字元：{value}')
    return value


def _write_log(path, message):
    try:
        directory = os.path.dirname(path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        with open(path, 'a', encoding='utf-8') as log_file:
            log_file.write(message + '\n')
    except Exception as error:
        print(f'寫入日誌檔失敗: {error}')


def log_and_display(message, failure=False):
    """一般事件寫入 system_log；失敗事件寫入 FIAL_system_log。"""
    full_message = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    _write_log(FAIL_LOG_FILENAME if failure else LOG_FILENAME, full_message)
    if monitor_text_area is not None:
        try:
            monitor_text_area.insert(tk.END, full_message + '\n')
            monitor_text_area.see(tk.END)
        except Exception:
            pass


def move_to_fail_folder(file_name, reason):
    """將失敗檔案搬到 FailFolder，並將搬移結果記錄到 FIAL log。"""
    os.makedirs(FAIL_DIR, exist_ok=True)
    source = os.path.join(SCAN_DIR, file_name)
    destination = os.path.join(FAIL_DIR, file_name)
    if os.path.exists(destination):
        base, extension = os.path.splitext(file_name)
        destination = os.path.join(
            FAIL_DIR,
            f'{base}_{datetime.now().strftime("%H%M%S")}{extension}',
        )
    try:
        if os.path.exists(source):
            shutil.move(source, destination)
        log_and_display(
            f'[失敗轉移][檔案={file_name}][原因={reason}][目錄={FAIL_DIR}]',
            failure=True,
        )
        return True
    except Exception as error:
        log_and_display(
            f'[失敗轉移失敗][檔案={file_name}][原因={reason}][錯誤={error}]',
            failure=True,
        )
        return False


# ==============================================================================
# Excel 資料處理
# ==============================================================================
def normalize_column_name(column_name):
    text = str(column_name or '').strip().lower().replace(' ', '')
    text = text.replace('（', '(').replace('）', ')')
    return re.sub(r'[^0-9a-z\u4e00-\u9fff]+', '', text)


def align_dataframe_columns(df):
    if df is None:
        return df
    df = df.copy()
    df.columns = [str(column).strip() for column in df.columns]
    aliases = {
        '工单号': '工單號', '工單號': '工單號', '工单': '工單號',
        '工號': '工單號', 'workorder': '工單號', 'orderno': '工單號', 'job': '工單號',
        'panelno': 'PANEL_NO', 'panel': 'PANEL_NO', 'panelnumber': 'PANEL_NO',
        'panelnum': 'PANEL_NO', 'panel_no': 'PANEL_NO', 'pnl': 'PANEL_NO',
        '併板主pnl': '併板主PNL', '主pnl': '併板主PNL', 'pnlmain': '併板主PNL',
        '序號': '序號(SN)', '序號sn': '序號(SN)', 'sn': '序號(SN)',
        's/n': '序號(SN)', 'serialnumber': '序號(SN)',
        '序號s1sn': '序號(S1SN)', 's1sn': '序號(S1SN)',
        '序號s2sn': '序號(S2SN)', 's2sn': '序號(S2SN)',
    }
    aliases = {normalize_column_name(k): v for k, v in aliases.items()}
    rename_map = {}
    for column in df.columns:
        target = aliases.get(normalize_column_name(column))
        if not target:
            continue
        if target in df.columns and target != column:
            df[target] = df[target].combine_first(df[column])
        else:
            rename_map[column] = target
    return df.rename(columns=rename_map) if rename_map else df


def load_all_excel_files():
    output_columns = ['工單號', 'PANEL_NO', '序號(SN)', '併板主PNL', '序號(S1SN)', '序號(S2SN)']
    data_frames = []
    for file_path in glob.glob(os.path.join('.', '*.xlsx')):
        filename = os.path.basename(file_path)
        try:
            df = align_dataframe_columns(pd.read_excel(file_path, dtype={'PANEL_NO': str}))
            if 'PANEL_NO' not in df.columns and '併板主PNL' in df.columns:
                df['PANEL_NO'] = df['併板主PNL']
            if '併板主PNL' not in df.columns and 'PANEL_NO' in df.columns:
                df['併板主PNL'] = df['PANEL_NO']
            if '序號(SN)' not in df.columns:
                source = '序號(S1SN)' if '序號(S1SN)' in df.columns else '序號(S2SN)'
                if source in df.columns:
                    df['序號(SN)'] = df[source]
            for column in ('序號(S1SN)', '序號(S2SN)'):
                if column not in df.columns and '序號(SN)' in df.columns:
                    df[column] = df['序號(SN)']
            if '工單號' not in df.columns:
                print(f'⚠️ 檔案 [{filename}] 缺少必須欄位 [工單號]，將跳過此檔案。')
                continue
            filtered = df.reindex(columns=output_columns).copy()
            for column in output_columns:
                filtered[column] = filtered[column].apply(lambda value: value.strip() if isinstance(value, str) else value)
            filtered['來源檔案'] = filename
            data_frames.append(filtered)
        except Exception as error:
            print(f'❌ 讀取檔案失敗 [{filename}]: {error}')
    return pd.concat(data_frames, ignore_index=True) if data_frames else pd.DataFrame(columns=output_columns)


def find_sn_by_panel_df(target_panel, buffer, bsn):
    if buffer is None or buffer.empty or 'PANEL_NO' not in buffer.columns:
        return pd.DataFrame()
    target = str(target_panel).strip().upper()
    result = buffer[buffer['PANEL_NO'].astype(str).str.strip().str.upper() == target].copy()
    if result.empty:
        return pd.DataFrame()
    rows = []
    for _, row in result.iterrows():
        value = row.get('序號(S1SN)', row.get('序號(SN)', ''))
        if pd.notna(value):
            item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
    for _, row in result.iterrows():
        value = row.get('序號(S2SN)')
        if pd.notna(value):
            item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
    try:
        index = int(bsn)
    except (TypeError, ValueError):
        return pd.DataFrame()
    return pd.DataFrame(rows).iloc[[index - 1]].copy() if 0 < index <= len(rows) else pd.DataFrame()


# ==============================================================================
# MySQL 上拋與二次確認
# ==============================================================================
def upload_to_mysql(sn_param, file_datetime, side, status, cfg, mysql_job, panel_no, file_name, model_name):
    if cfg['setting'].getint('MySQL_FLAG', 0) != 1:
        log_and_display('MySQL_FLAG 未啟用，跳過上拋。', failure=True)
        return False
    try:
        mysql_type = mysql_safe_identifier(cfg['setting'].get('MySQL_TYPE', 'CARD'), 'MySQL_TYPE')
        base_table = mysql_safe_identifier(cfg['setting'].get('TableDetailStr', 'STA1'), 'TableDetailStr')
        detail_table = f'{base_table}_{"B" if side == "B" else "T"}'
    except ValueError as error:
        log_and_display(f'MySQL 表名設定錯誤: {error}', failure=True)
        return False

    for attempt in range(1, 4):
        conn = None
        try:
            log_and_display(
                f'[重送上傳][SN={sn_param}][PANEL={panel_no}][Job={mysql_job}]'
                f'[DB寫入][第{attempt}/3次]' if attempt == 1 else
                f'[重送上傳][SN={sn_param}][PANEL={panel_no}][Job={mysql_job}]'
                f'[DB寫入][重試第{attempt}/3次]',
            )
            conn = pymysql.connect(
                host=cfg['setting'].get('MySQL_ServerIP'),
                user=cfg['setting'].get('MySQL_username'),
                password=cfg['setting'].get('MySQL_Password'),
                database=cfg['setting'].get('MySQL_DB'),
                charset='utf8', connect_timeout=5,
            )
            with conn.cursor() as cursor:
                if cfg['setting'].getint('MySQL_InsertFlag', 0) == 1:
                    cursor.execute(
                        f'INSERT INTO `{mysql_type}` SET iSN=%s, SMT_PN=%s, PANEL_SN=%s, `{detail_table}`=101 '
                        'ON DUPLICATE KEY UPDATE SMT_PN=%s, PANEL_SN=%s',
                        (sn_param, mysql_job, panel_no, mysql_job, panel_no),
                    )
                else:
                    cursor.execute(
                        f'UPDATE `{mysql_type}` SET SMT_PN=%s, PANEL_SN=%s, `{detail_table}`=101 WHERE iSN=%s',
                        (mysql_job, panel_no, sn_param),
                    )
                cursor.execute(
                    f'INSERT INTO `{detail_table}` SET iSN=%s, errorCode=%s, JobNum=%s, ModelName=%s, '
                    'operator=%s, Station=%s, StartTime=%s, StopTime=%s, logfilename=%s, log=%s',
                    (sn_param, status, mysql_job, model_name, cfg['setting'].get('MySQL_Operator'),
                     cfg['setting'].get('MySQL_Station'), file_datetime, file_datetime, file_name, file_name),
                )
            conn.commit()
            with conn.cursor() as cursor:
                cursor.execute(f'SELECT `{detail_table}` FROM `{mysql_type}` WHERE iSN=%s', (sn_param,))
                main_ok = (cursor.fetchone() or [None])[0] == 101
                cursor.execute(
                    f'SELECT 1 FROM `{detail_table}` WHERE iSN=%s AND JobNum=%s AND logfilename=%s LIMIT 1',
                    (sn_param, mysql_job, file_name),
                )
                detail_ok = cursor.fetchone() is not None
            if main_ok and detail_ok:
                log_and_display(f'[二次確認成功][SN={sn_param}][PANEL={panel_no}]')
                return True
            log_and_display(
                f'[二次確認失敗][SN={sn_param}][PANEL={panel_no}]'
                f'[主表={main_ok}][明細表={detail_ok}]', failure=True,
            )
            return False
        except Exception as error:
            log_and_display(
                f'[重送上傳][SN={sn_param}][DB寫入][失敗][第{attempt}/3次][錯誤={error}]',
                failure=True,
            )
            if attempt < 3:
                time.sleep(2)
            else:
                log_and_display(f'[重試上限達成][SN={sn_param}]', failure=True)
                return False
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
    return False


def scan_folder_loop():
    while True:
        try:
            current_cfg = load_config()
            files = [name for name in os.listdir(SCAN_DIR) if os.path.isfile(os.path.join(SCAN_DIR, name))]
            if files:
                log_and_display(f'偵測到 {len(files)} 個新檔案，開始比對並上拋...')
            for file_name in files:
                file_path = os.path.join(SCAN_DIR, file_name)
                try:
                    parts = file_name.split('_')
                    if len(parts) < 5:
                        log_and_display(f'⚠️ 檔案 [{file_name}] 名稱格式不符，跳過處理。', failure=True)
                        move_to_fail_folder(file_name, '檔名格式不符')
                        continue
                    panel_no, status, bsn = parts[2].strip(), parts[3], parts[4]
                    model_name = parts[5][:-1]
                    side = parts[-2][-1]
                    time_part = os.path.splitext(parts[-1])[0].strip().upper()
                    if time_part in ('', 'NONE'):
                        file_dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        try:
                            file_dt = datetime.strptime(time_part, '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
                        except ValueError:
                            file_dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            log_and_display(f'⚠️ 檔案 [{file_name}] 時間格式錯誤。', failure=True)
                    if status.upper() == 'FAIL':
                        log_and_display(f'ℹ️ 檔案 [{file_name}] 狀態為 FAIL，跳過上拋資料庫。', failure=True)
                        move_to_fail_folder(file_name, '狀態為 FAIL')
                        continue
                    results = find_sn_by_panel_df(panel_no, data_buffer, bsn)
                    if results.empty:
                        log_and_display(f'⚠️ PANEL_NO [{panel_no}] 比對失敗。', failure=True)
                        move_to_fail_folder(file_name, 'PANEL_NO 比對失敗')
                        continue
                    log_and_display(f'PANEL_NO [{panel_no}] 比對成功，找到 {len(results)} 筆 SN 資料，開始逐筆上拋...')
                    upload_results = []
                    for _, row in results.iterrows():
                        actual_sn = str(row['序號(SN)']).strip()
                        upload_results.append(upload_to_mysql(actual_sn, file_dt, side, status, current_cfg, parts[1], panel_no, file_name, model_name))
                    if not all(upload_results):
                        log_and_display(f'檔案 [{file_name}] 上拋失敗，移至 FailFolder。', failure=True)
                        move_to_fail_folder(file_name, '一筆或多筆 SN 上拋失敗')
                        continue
                    log_and_display(f'PANEL [{panel_no}] 全部 SN 上拋並二次確認成功。')
                except Exception as error:
                    log_and_display(f'⚠️ 處理個別檔案 {file_name} 時發生異常: {error}', failure=True)
                    move_to_fail_folder(file_name, f'處理異常: {error}')
                    continue
                if os.path.exists(file_path):
                    destination = os.path.join(BACKUP_DIR, file_name)
                    if os.path.exists(destination):
                        base, extension = os.path.splitext(file_name)
                        destination = os.path.join(BACKUP_DIR, f'{base}_{datetime.now().strftime("%H%M%S")}{extension}')
                    shutil.move(file_path, destination)
        except Exception as error:
            log_and_display(f'⚠️ 掃描迴圈發生嚴重錯誤: {error}', failure=True)
        time.sleep(30)


def open_monitor_ui():
    global monitor_text_area
    root = tk.Tk()
    root.title('系統執行監控中')
    root.geometry('550x380')
    root.resizable(False, False)
    tk.Label(root, text='🚀 系統自動監控服務中', font=('微軟正黑體', 12, 'bold'), fg='green').pack(pady=10)
    monitor_text_area = scrolledtext.ScrolledText(root, width=65, height=14, font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4')
    monitor_text_area.pack(padx=15, pady=5)
    log_and_display('監控服務啟動成功。')
    tk.Button(root, text='EXIT (停止並離開系統)', bg='#333333', fg='white', width=22, command=root.destroy).pack(pady=15)
    threading.Thread(target=scan_folder_loop, daemon=True).start()
    root.mainloop()


def create_setup_ui():
    global data_buffer
    config_data = load_config()
    setting = config_data['setting'] if 'setting' in config_data else {}
    root = tk.Tk()
    root.title('MySQL 參數設定調整')
    root.geometry('400x260')
    root.resizable(False, False)
    entries = {}
    keys = ['MySQL_Job', 'MySQL_ModelName', 'MySQL_Operator', 'MySQL_Station']
    for row, key in enumerate(keys):
        tk.Label(root, text=f'{key}:', font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky='e', padx=15, pady=6)
        entry = tk.Entry(root, width=30)
        entry.grid(row=row, column=1, padx=15, pady=6)
        entry.insert(0, setting.get(key, ''))
        entries[key] = entry
    def confirm():
        if save_config(config_data, *(entries[key].get().strip() for key in keys)):
            root.destroy(); open_monitor_ui()
    def cancel():
        root.destroy(); open_monitor_ui()
    frame = tk.Frame(root); frame.grid(row=4, column=0, columnspan=2, pady=15)
    tk.Button(frame, text='否 (不存檔直接監控)', bg='#d9534f', fg='white', width=16, command=cancel).pack(side='left', padx=15)
    tk.Button(frame, text='確認 (存檔並監控)', bg='#5cb85c', fg='white', width=16, command=confirm).pack(side='right', padx=15)
    root.mainloop()


if __name__ == '__main__':
    data_buffer = load_all_excel_files()
    try:
        import pymysql
    except ImportError:
        root = tk.Tk(); root.title('警告'); root.geometry('300x100')
        tk.Label(root, text='警告：您未安裝 pymysql 模組，無法繼續執行。', fg='red').pack(pady=20)
        root.mainloop(); raise SystemExit(1)
    create_setup_ui()
