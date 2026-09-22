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
CONFIG_FILENAME = 'config.ini'
_DATE = datetime.now().strftime('%Y%m%d')
_DEFAULT_PATHS = {
    'SCAN_DIR': './ScanFolder',
    'BACKUP_DIR': './BackupFolder',
    'FAIL_DIR': './FailFolder',
    'LOG_FILENAME': f'./LOG/system_log{_DATE}.txt',
    'FAIL_LOG_FILENAME': f'./LOG/Fail_system_log{_DATE}.txt',
}


def _load_path_config():
    """讀取並修復路徑設定，包含舊版 FIAL_system_log 拼字的相容處理。"""
    parser = configparser.ConfigParser()
    changed = False
    if os.path.exists(CONFIG_FILENAME):
        parser.read(CONFIG_FILENAME, encoding='utf-8')
    if not parser.has_section('PATH'):
        parser['PATH'] = {}
        changed = True

    section = parser['PATH']
    for key, default in _DEFAULT_PATHS.items():
        if not section.get(key, '').strip():
            section[key] = default
            changed = True

    # 舊版曾使用 FIAL_system_log；自動改成正確且固定的 Fail_system_log。
    fail_log = section.get('FAIL_LOG_FILENAME', '').strip()
    if not fail_log or 'FIAL_system_log' in os.path.basename(fail_log):
        fail_log = fail_log.replace('FIAL_system_log', 'Fail_system_log') or _DEFAULT_PATHS['FAIL_LOG_FILENAME']
        section['FAIL_LOG_FILENAME'] = fail_log
        changed = True

    if changed or not os.path.exists(CONFIG_FILENAME):
        with open(CONFIG_FILENAME, 'w', encoding='utf-8') as config_file:
            parser.write(config_file)
    return section


path_config = _load_path_config()
SCAN_DIR = path_config.get('SCAN_DIR', _DEFAULT_PATHS['SCAN_DIR']).strip()
BACKUP_DIR = path_config.get('BACKUP_DIR', _DEFAULT_PATHS['BACKUP_DIR']).strip()
FAIL_DIR = path_config.get('FAIL_DIR', _DEFAULT_PATHS['FAIL_DIR']).strip()
LOG_FILENAME = path_config.get('LOG_FILENAME', _DEFAULT_PATHS['LOG_FILENAME']).strip()
FAIL_LOG_FILENAME = path_config.get(
    'FAIL_LOG_FILENAME', _DEFAULT_PATHS['FAIL_LOG_FILENAME']
).strip()

for directory in (
    SCAN_DIR,
    BACKUP_DIR,
    FAIL_DIR,
    os.path.dirname(LOG_FILENAME),
    os.path.dirname(FAIL_LOG_FILENAME),
):
    if directory:
        os.makedirs(directory, exist_ok=True)

monitor_text_area = None

if not os.path.exists(INI_FILENAME):
    with open(INI_FILENAME, 'w', encoding='utf-8') as config_file:
        config_file.write('''[SystemID]
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
''')


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
    del config_data
    targets = {
        'MySQL_Job': job,
        'MySQL_ModelName': model,
        'MySQL_Operator': operator,
        'MySQL_Station': station,
    }
    try:
        with open(INI_FILENAME, 'r', encoding='utf-8') as config_file:
            lines = config_file.readlines()
        output = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith(('/', '#', ';')) or '=' not in stripped:
                output.append(line)
                continue
            key = stripped.split('=', 1)[0].strip()
            output.append(f'{key}={targets[key]}\n' if key in targets else line)
        with open(INI_FILENAME, 'w', encoding='utf-8') as config_file:
            config_file.writelines(output)
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
        print(f'寫入日誌檔失敗 [{path}]: {error}')


def log_and_display(message, failure=False):
    """正常事件寫 system_log；失敗事件只寫 Fail_system_log。"""
    full_message = f'[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}'
    _write_log(FAIL_LOG_FILENAME if failure else LOG_FILENAME, full_message)
    if monitor_text_area is not None:
        try:
            monitor_text_area.insert(tk.END, full_message + '\n')
            monitor_text_area.see(tk.END)
        except Exception:
            pass


def _unique_destination(directory, file_name):
    destination = os.path.join(directory, file_name)
    if not os.path.exists(destination):
        return destination
    base, extension = os.path.splitext(file_name)
    return os.path.join(directory, f'{base}_{datetime.now():%Y%m%d_%H%M%S_%f}{extension}')


def move_to_fail_folder(file_name, reason):
    """將失敗檔案可靠地移到 FailFolder，並記錄搬移結果。"""
    try:
        os.makedirs(FAIL_DIR, exist_ok=True)
        source = os.path.join(SCAN_DIR, file_name)
        if not os.path.isfile(source):
            log_and_display(
                f'[失敗轉移失敗][檔案={file_name}][原因={reason}]'
                f'[找不到來源={source}]', failure=True,
            )
            return False
        destination = _unique_destination(FAIL_DIR, file_name)
        shutil.move(source, destination)
        log_and_display(
            f'[失敗轉移成功][檔案={file_name}][原因={reason}]'
            f'[目錄={destination}]', failure=True,
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
    columns = ['工單號', 'PANEL_NO', '序號(SN)', '併板主PNL', '序號(S1SN)', '序號(S2SN)']
    frames = []
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
            filtered = df.reindex(columns=columns).copy()
            for column in columns:
                filtered[column] = filtered[column].apply(
                    lambda value: value.strip() if isinstance(value, str) else value
                )
            filtered['來源檔案'] = filename
            frames.append(filtered)
        except Exception as error:
            print(f'❌ 讀取檔案失敗 [{filename}]: {error}')
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=columns)


def find_sn_by_panel_df(target_panel, buffer, bsn):
    if buffer is None or buffer.empty or 'PANEL_NO' not in buffer.columns:
        return pd.DataFrame()
    target = str(target_panel).strip().upper()
    result = buffer[buffer['PANEL_NO'].astype(str).str.strip().str.upper() == target]
    rows = []
    for _, row in result.iterrows():
        value = row.get('序號(S1SN)', row.get('序號(SN)', ''))
        if pd.notna(value) and str(value).strip():
            item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
    for _, row in result.iterrows():
        value = row.get('序號(S2SN)')
        if pd.notna(value) and str(value).strip():
            item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
    try:
        index = int(str(bsn).strip())
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
        side = str(side).strip().upper()
        if side not in ('T', 'B'):
            raise ValueError(f'Side 僅允許 T 或 B：{side}')
        detail_table = f'{base_table}_{side}'
    except ValueError as error:
        log_and_display(f'MySQL 表名設定錯誤: {error}', failure=True)
        return False

    for attempt in range(1, 4):
        conn = None
        try:
            log_and_display(f'[DB寫入][SN={sn_param}][PANEL={panel_no}][第{attempt}/3次]')
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
                main_value = (cursor.fetchone() or [None])[0]
                main_ok = str(main_value) == '101'
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
                f'[主表={main_value}][明細表={detail_ok}]', failure=True,
            )
            return False
        except Exception as error:
            log_and_display(f'[DB寫入失敗][SN={sn_param}][第{attempt}/3次][錯誤={error}]', failure=True)
            if attempt < 3:
                time.sleep(2)
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
    log_and_display(f'[重試上限達成][SN={sn_param}]', failure=True)
    return False


def scan_folder_loop():
    while True:
        try:
            current_cfg = load_config()
            files = [name for name in os.listdir(SCAN_DIR)
                     if os.path.isfile(os.path.join(SCAN_DIR, name))]
            if files:
                log_and_display(f'偵測到 {len(files)} 個新檔案，開始比對並上拋...')
            for file_name in files:
                file_path = os.path.join(SCAN_DIR, file_name)
                succeeded = False
                try:
                    parts = file_name.split('_')
                    # 後續會使用 parts[5]，所以至少需要 6 段。
                    if len(parts) < 6:
                        raise ValueError('檔名格式不符，至少需要 6 個底線分隔欄位')
                    panel_no, status, bsn = parts[2].strip(), parts[3].strip(), parts[4].strip()
                    model_name = parts[5][:-1]
                    side = parts[-2][-1:].upper()
                    time_part = os.path.splitext(parts[-1])[0].strip().upper()
                    if time_part in ('', 'NONE'):
                        file_dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        try:
                            file_dt = datetime.strptime(time_part, '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
                        except ValueError:
                            file_dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            log_and_display(f'⚠️ 檔案 [{file_name}] 時間格式錯誤，改用目前時間。', failure=True)
                    if status.upper() == 'FAIL':
                        log_and_display(f'檔案 [{file_name}] 狀態為 FAIL，跳過上拋。', failure=True)
                        move_to_fail_folder(file_name, '狀態為 FAIL')
                        continue
                    results = find_sn_by_panel_df(panel_no, data_buffer, bsn)
                    if results.empty:
                        log_and_display(f'PANEL_NO [{panel_no}] 比對失敗。', failure=True)
                        move_to_fail_folder(file_name, 'PANEL_NO 比對失敗')
                        continue
                    upload_results = [
                        upload_to_mysql(str(row['序號(SN)']).strip(), file_dt, side, status,
                                        current_cfg, parts[1], panel_no, file_name, model_name)
                        for _, row in results.iterrows()
                    ]
                    if not upload_results or not all(upload_results):
                        log_and_display(f'檔案 [{file_name}] 上拋失敗，移至 FailFolder。', failure=True)
                        move_to_fail_folder(file_name, '一筆或多筆 SN 上拋失敗')
                        continue
                    succeeded = True
                    log_and_display(f'檔案 [{file_name}] 全部 SN 上拋並二次確認成功。')
                except Exception as error:
                    log_and_display(f'⚠️ 處理個別檔案 {file_name} 時發生異常: {error}', failure=True)
                    move_to_fail_folder(file_name, f'處理異常: {error}')
                if succeeded and os.path.exists(file_path):
                    try:
                        shutil.move(file_path, _unique_destination(BACKUP_DIR, file_name))
                    except Exception as error:
                        log_and_display(f'成功檔案搬移至 BackupFolder 失敗: {error}', failure=True)
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
