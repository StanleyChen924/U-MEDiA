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
import sys
import pandas as pd

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Use paths relative to this script, not the process working directory. This is
# important when the program is started by a shortcut or packaged as an EXE.
# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INI_FILENAME = os.path.join(BASE_DIR, 'MySQLConfig.ini')
CONFIG_FILENAME = os.path.join(BASE_DIR, 'config.ini')
TODAY = datetime.now().strftime('%Y%m%d')
DEFAULT_PATHS = {
    'SCAN_DIR': 'ScanFolder',
    'BACKUP_DIR': 'BackupFolder',
    'FAIL_DIR': 'FailFolder',
    'LOG_FILENAME': os.path.join('LOG', f'system_log{TODAY}.txt'),
    'FAIL_LOG_FILENAME': os.path.join('LOG', f'Fail_system_log{TODAY}.txt'),
}


def _absolute_path(value):
    value = os.path.expandvars(os.path.expanduser(str(value or '').strip()))
    return value if os.path.isabs(value) else os.path.join(BASE_DIR, value)


def _load_paths():
    parser = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILENAME):
        parser.read(CONFIG_FILENAME, encoding='utf-8')
    if not parser.has_section('PATH'):
        parser['PATH'] = {}
    section = parser['PATH']
    changed = False
    for key, default in DEFAULT_PATHS.items():
        if not section.get(key, '').strip():
            section[key] = default
            changed = True
    # Migrate the old typo and never use it again.
    old = section.get('FAIL_LOG_FILENAME', '')
    if 'FIAL_system_log' in old:
        section['FAIL_LOG_FILENAME'] = old.replace('FIAL_system_log', 'Fail_system_log')
        changed = True
    if changed or not os.path.exists(CONFIG_FILENAME):
        with open(CONFIG_FILENAME, 'w', encoding='utf-8') as stream:
            parser.write(stream)
    return section


paths = _load_paths()
SCAN_DIR = _absolute_path(paths.get('SCAN_DIR', DEFAULT_PATHS['SCAN_DIR']))
BACKUP_DIR = _absolute_path(paths.get('BACKUP_DIR', DEFAULT_PATHS['BACKUP_DIR']))
FAIL_DIR = _absolute_path(paths.get('FAIL_DIR', DEFAULT_PATHS['FAIL_DIR']))
LOG_FILENAME = _absolute_path(paths.get('LOG_FILENAME', DEFAULT_PATHS['LOG_FILENAME']))
FAIL_LOG_FILENAME = _absolute_path(paths.get('FAIL_LOG_FILENAME', DEFAULT_PATHS['FAIL_LOG_FILENAME']))
for directory in (SCAN_DIR, BACKUP_DIR, FAIL_DIR, os.path.dirname(LOG_FILENAME), os.path.dirname(FAIL_LOG_FILENAME)):
    os.makedirs(directory, exist_ok=True)

monitor_text_area = None

if not os.path.exists(INI_FILENAME):
    with open(INI_FILENAME, 'w', encoding='utf-8') as stream:
        stream.write('''[SystemID]\nDevice_ID=0x0013\nVendor_ID=0x168C\nSSYS_ID=0x2051\nSSYS_VEND_ID=0x168C\n\n[setting]\nMySQL_FLAG=1\nMySQL_InsertFlag=1\nMySQL_BeforStation=and STA1 > 100\nTableDetailStr=STA1\nMySQL_TYPE=CARD\n''')


def load_config():
    result = configparser.ConfigParser(comment_prefixes=('#', ';', '//'), inline_comment_prefixes=('#', ';', '//'), strict=False)
    result.optionxform = str
    result.read(INI_FILENAME, encoding='utf-8')
    return result


def save_config(config_data, job, model, operator, station):
    del config_data
    targets = {'MySQL_Job': job, 'MySQL_ModelName': model, 'MySQL_Operator': operator, 'MySQL_Station': station}
    try:
        with open(INI_FILENAME, encoding='utf-8') as stream:
            lines = stream.readlines()
        output = []
        for line in lines:
            stripped = line.strip()
            if stripped.startswith(('/', '#', ';')) or '=' not in stripped:
                output.append(line)
                continue
            key = stripped.split('=', 1)[0].strip()
            output.append(f'{key}={targets[key]}\n' if key in targets else line)
        with open(INI_FILENAME, 'w', encoding='utf-8') as stream:
            stream.writelines(output)
        return True
    except Exception as error:
        messagebox.showerror('錯誤', f'存檔失敗：{error}')
        return False


def _write_log(path, message):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'a', encoding='utf-8') as stream:
            stream.write(message + '\n')
    except Exception as error:
        print(f'寫入日誌檔失敗 [{path}]: {error}')


def log_and_display(message, failure=False):
    full = f'[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}'
    _write_log(FAIL_LOG_FILENAME if failure else LOG_FILENAME, full)
    if monitor_text_area is not None:
        try:
            monitor_text_area.insert(tk.END, full + '\n')
            monitor_text_area.see(tk.END)
        except Exception:
            pass


def _unique_path(directory, name):
    path = os.path.join(directory, name)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(name)
    return os.path.join(directory, f'{stem}_{datetime.now():%Y%m%d_%H%M%S_%f}{ext}')


def move_to_fail_folder(file_name, reason):
    """Move a failed input file and report both success and failure explicitly."""
    source = os.path.abspath(os.path.join(SCAN_DIR, file_name))
    try:
        os.makedirs(FAIL_DIR, exist_ok=True)
        if not os.path.isfile(source):
            log_and_display(f'[失敗轉移失敗][檔案={file_name}][原因={reason}][來源不存在={source}]', failure=True)
            return False
        destination = _unique_path(FAIL_DIR, os.path.basename(file_name))
        shutil.move(source, destination)
        if os.path.exists(source) or not os.path.exists(destination):
            raise OSError(f'搬移後檔案驗證失敗: {destination}')
        log_and_display(f'[失敗轉移成功][檔案={file_name}][原因={reason}][來源={source}][目標={destination}]', failure=True)
        return True
    except Exception as error:
        log_and_display(f'[失敗轉移失敗][檔案={file_name}][原因={reason}][錯誤={error}]', failure=True)
        return False


def mysql_safe_identifier(value, field_name='identifier'):
    value = str(value or '').strip()
    if not re.fullmatch(r'[A-Za-z0-9_]+', value):
        raise ValueError(f'{field_name} 含有非法字元：{value}')
    return value


def normalize_column_name(value):
    text = str(value or '').strip().lower().replace(' ', '')
    text = text.replace('（', '(').replace('）', ')')
    return re.sub(r'[^0-9a-z\u4e00-\u9fff]+', '', text)


def align_dataframe_columns(df):
    df = df.copy()
    df.columns = [str(column).strip() for column in df.columns]
    aliases = {'工单号':'工單號','工單號':'工單號','工单':'工單號','工號':'工單號','workorder':'工單號','orderno':'工單號','job':'工單號','panelno':'PANEL_NO','panel':'PANEL_NO','序號(SN)':'序號(SN)','併板主PNL':'併板主PNL','序號(S1SN)':'序號(S1SN)','序號(S2SN)':'序號(S2SN)'}
    aliases = {normalize_column_name(k): v for k, v in aliases.items()}
    rename = {}
    for column in df.columns:
        target = aliases.get(normalize_column_name(column))
        if target:
            if target in df.columns and target != column:
                df[target] = df[target].combine_first(df[column])
            else:
                rename[column] = target
    return df.rename(columns=rename)


def load_all_excel_files():
    columns = ['工單號', 'PANEL_NO', '序號(SN)', '併板主PNL', '序號(S1SN)', '序號(S2SN)']
    two_column_panel_sources = {'序號(SN)', '併板主PNL', '序號(S1SN)', '序號(S2SN)'}
    frames = []
    for path in glob.glob(os.path.join(BASE_DIR, '*.xlsx')):
        name = os.path.basename(path)
        try:
            df = align_dataframe_columns(pd.read_excel(path, dtype={'PANEL_NO': str}))
            # Some Excel files contain exactly two columns and omit PANEL_NO.
            # In that format, the second column is the PANEL_NO lookup column,
            # regardless of whether it is named SN, PNL, S1SN, or S2SN.
            if (
                'PANEL_NO' not in df
                and len(df.columns) == 2
                and df.columns[1] in two_column_panel_sources
            ):
                df['PANEL_NO'] = df.iloc[:, 1]
            if 'PANEL_NO' not in df and '併板主PNL' in df: df['PANEL_NO'] = df['併板主PNL']
            if '併板主PNL' not in df and 'PANEL_NO' in df: df['併板主PNL'] = df['PANEL_NO']
            if '序號(SN)' not in df:
                for source in ('序號(S1SN)', '序號(S2SN)'):
                    if source in df: df['序號(SN)'] = df[source]; break
            for column in ('序號(S1SN)', '序號(S2SN)'):
                if column not in df and '序號(SN)' in df: df[column] = df['序號(SN)']
            if '工單號' not in df:
                print(f'⚠️ 檔案 [{name}] 缺少必須欄位 [工單號]，將跳過此檔案。'); continue
            filtered = df.reindex(columns=columns).copy()
            for column in columns:
                filtered[column] = filtered[column].apply(lambda value: value.strip() if isinstance(value, str) else value)
            frames.append(filtered)
        except Exception as error:
            print(f'❌ 讀取檔案失敗 [{name}]: {error}')
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=columns)


def find_sn_by_panel_df(panel, buffer, bsn):
    if buffer is None or buffer.empty or 'PANEL_NO' not in buffer: return pd.DataFrame()
    result = buffer[buffer['PANEL_NO'].astype(str).str.strip().str.upper() == str(panel).strip().upper()]
    rows = []
    for _, row in result.iterrows():
        for column in ('序號(S1SN)', '序號(S2SN)'):
            value = row.get(column)
            if pd.notna(value) and str(value).strip():
                item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
        if not any(pd.notna(row.get(column)) and str(row.get(column)).strip() for column in ('序號(S1SN)', '序號(S2SN)')):
            value = row.get('序號(SN)')
            if pd.notna(value) and str(value).strip():
                item = row.copy(); item['序號(SN)'] = str(value).strip(); rows.append(item)
    try: index = int(str(bsn).strip())
    except (TypeError, ValueError): return pd.DataFrame()
    return pd.DataFrame(rows).iloc[[index - 1]].copy() if 0 < index <= len(rows) else pd.DataFrame()


def upload_to_mysql(sn, file_dt, side, status, cfg, job, panel, file_name, model):
    if cfg['setting'].getint('MySQL_FLAG', 0) != 1:
        log_and_display('MySQL_FLAG 未啟用，跳過上拋。', failure=True); return False
    try:
        table = mysql_safe_identifier(cfg['setting'].get('MySQL_TYPE', 'CARD'), 'MySQL_TYPE')
        base = mysql_safe_identifier(cfg['setting'].get('TableDetailStr', 'STA1'), 'TableDetailStr')
        side = str(side).strip().upper()
        if side not in ('T', 'B'): raise ValueError(f'Side 僅允許 T 或 B：{side}')
        detail = f'{base}_{side}'
    except ValueError as error:
        log_and_display(f'MySQL 表名設定錯誤: {error}', failure=True); return False
    for attempt in range(1, 4):
        conn = None
        try:
            import pymysql
            conn = pymysql.connect(host=cfg['setting'].get('MySQL_ServerIP'), user=cfg['setting'].get('MySQL_username'), password=cfg['setting'].get('MySQL_Password'), database=cfg['setting'].get('MySQL_Database'), port=cfg['setting'].getint('MySQL_Port', 3306), charset='utf8', autocommit=False)
            with conn.cursor() as cursor:
                if cfg['setting'].getint('MySQL_InsertFlag', 0) == 1:
                    cursor.execute(f'INSERT INTO `{table}` SET iSN=%s, SMT_PN=%s, PANEL_SN=%s, `{detail}`=101 ON DUPLICATE KEY UPDATE SMT_PN=%s, PANEL_SN=%s', (sn, job, panel, job, panel))
                else:
                    cursor.execute(f'UPDATE `{table}` SET SMT_PN=%s, PANEL_SN=%s, `{detail}`=101 WHERE iSN=%s', (job, panel, sn))
                cursor.execute(f'INSERT INTO `{detail}` SET iSN=%s, errorCode=%s, JobNum=%s, ModelName=%s, operator=%s, Station=%s, StartTime=%s, StopTime=%s, logfilename=%s, log=%s', (sn, status, job, model, cfg['setting'].get('MySQL_Operator'), cfg['setting'].get('MySQL_Station'), file_dt, file_dt, file_name, status))
            conn.commit()
            with conn.cursor() as cursor:
                cursor.execute(f'SELECT `{detail}` FROM `{table}` WHERE iSN=%s', (sn,)); main_ok = str((cursor.fetchone() or [None])[0]) == '101'
                cursor.execute(f'SELECT 1 FROM `{detail}` WHERE iSN=%s AND JobNum=%s AND logfilename=%s LIMIT 1', (sn, job, file_name)); detail_ok = cursor.fetchone() is not None
            if main_ok and detail_ok: return True
            log_and_display(f'[二次確認失敗][SN={sn}][PANEL={panel}][主表={main_ok}][明細表={detail_ok}]', failure=True); return False
        except Exception as error:
            log_and_display(f'[DB寫入失敗][SN={sn}][第{attempt}/3次][錯誤={error}]', failure=True)
            if attempt < 3: time.sleep(2)
        finally:
            if conn is not None:
                try: conn.close()
                except Exception: pass
    return False


def scan_folder_loop():
    log_and_display(f'掃描服務啟動，BASE_DIR={BASE_DIR}, SCAN_DIR={SCAN_DIR}')
    while True:
        try:
            cfg = load_config()
            for file_name in [n for n in os.listdir(SCAN_DIR) if os.path.isfile(os.path.join(SCAN_DIR, n))]:
                source = os.path.join(SCAN_DIR, file_name)
                succeeded = False
                try:
                    parts = file_name.split('_')
                    if len(parts) < 6: raise ValueError('檔名格式不符，至少需要 6 個欄位')
                    panel, status, bsn = parts[2].strip(), parts[3].strip(), parts[4].strip()
                    model, side = parts[5][:-1], parts[-2][-1:].upper()
                    stamp = os.path.splitext(parts[-1])[0].strip().upper()
                    try: file_dt = datetime.strptime(stamp, '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M:%S') if stamp and stamp != 'NONE' else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    except ValueError: file_dt = datetime.now().strftime('%Y-%m-%d %H:%M:%S'); log_and_display(f'時間格式錯誤: {file_name}', failure=True)
                    if status.upper() == 'FAIL': move_to_fail_folder(file_name, '狀態為 FAIL'); continue
                    results = find_sn_by_panel_df(panel, data_buffer, bsn)
                    if results.empty:
                        log_and_display(f'比對失敗：在 Excel 緩衝區中找不到 PANEL_NO [{panel}] 的任何資料，此檔案不上拋。', failure=True)
                        move_to_fail_folder(file_name, 'PANEL_NO 比對失敗'); continue
                    uploads = [upload_to_mysql(str(row['序號(SN)']).strip(), file_dt, side, status, cfg, parts[1], panel, file_name, model) for _, row in results.iterrows()]
                    if not uploads or not all(uploads):
                        log_and_display(f'檔案 [{file_name}] 上拋失敗，移至 FailFolder。', failure=True)
                        move_to_fail_folder(file_name, '一筆或多筆 SN 上拋失敗'); continue
                    succeeded = True
                except Exception as error:
                    log_and_display(f'處理個別檔案 {file_name} 發生異常: {error}', failure=True)
                    move_to_fail_folder(file_name, f'處理異常: {error}')
                if succeeded and os.path.isfile(source):
                    try: shutil.move(source, _unique_path(BACKUP_DIR, file_name))
                    except Exception as error: log_and_display(f'搬移至 BackupFolder 失敗: {error}', failure=True)
        except Exception as error:
            log_and_display(f'掃描迴圈發生嚴重錯誤: {error}', failure=True)
        time.sleep(30)


def open_monitor_ui():
    global monitor_text_area
    root = tk.Tk(); root.title('系統執行監控中'); root.geometry('550x380'); root.resizable(False, False)
    tk.Label(root, text='🚀 系統自動監控服務中', font=('微軟正黑體', 12, 'bold'), fg='green').pack(pady=10)
    monitor_text_area = scrolledtext.ScrolledText(root, width=65, height=14, font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4'); monitor_text_area.pack(padx=15, pady=5)
    log_and_display('監控服務啟動成功。')
    tk.Button(root, text='EXIT (停止並離開系統)', bg='#333333', fg='white', width=22, command=root.destroy).pack(pady=15)
    threading.Thread(target=scan_folder_loop, daemon=True).start(); root.mainloop()


def create_setup_ui():
    config_data = load_config(); setting = config_data['setting'] if 'setting' in config_data else {}
    root = tk.Tk(); root.title('MySQL 參數設定調整'); root.geometry('400x260'); root.resizable(False, False)
    entries = {}; keys = ['MySQL_Job', 'MySQL_ModelName', 'MySQL_Operator', 'MySQL_Station']
    for row, key in enumerate(keys):
        tk.Label(root, text=f'{key}:', font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky='e', padx=15, pady=6)
        entry = tk.Entry(root, width=30); entry.grid(row=row, column=1, padx=15, pady=6); entry.insert(0, setting.get(key, '')); entries[key] = entry
    def confirm():
        if save_config(config_data, *(entries[key].get().strip() for key in keys)): root.destroy(); open_monitor_ui()
    def cancel(): root.destroy(); open_monitor_ui()
    frame = tk.Frame(root); frame.grid(row=4, column=0, columnspan=2, pady=15)
    tk.Button(frame, text='否 (不存檔直接監控)', bg='#d9534f', fg='white', width=16, command=cancel).pack(side='left', padx=15)
    tk.Button(frame, text='確認 (存檔並監控)', bg='#5cb85c', fg='white', width=16, command=confirm).pack(side='right', padx=15)
    root.mainloop()


if __name__ == '__main__':
    data_buffer = load_all_excel_files()
    try:
        import pymysql
    except ImportError:
        root = tk.Tk(); root.title('警告'); root.geometry('300x100'); tk.Label(root, text='警告：您未安裝 pymysql 模組，無法繼續執行。', fg='red').pack(pady=20); root.mainloop(); raise SystemExit
    create_setup_ui()
