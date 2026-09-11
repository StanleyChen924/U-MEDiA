import os
import sys
import json
import datetime
import configparser
import subprocess
import tkinter as tk
from tkinter import messagebox
import winsound
import importlib
import pymysql


class CartonSystemApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Carton Label 掃描控制系統")
        self.root.tk.call("tk", "scaling", 1.3)
        self.root.geometry("780x585")
        self.app_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(self.app_dir)
        if getattr(sys, "frozen", False) and not os.path.exists(os.path.join(self.app_dir, "2544259S1S2.xlsx")) and os.path.exists(os.path.join(parent_dir, "2544259S1S2.xlsx")):
            self.app_dir = parent_dir
        
        # 1. 讀取並防呆設定檔
        self.load_configs()

        self.log_dir = os.path.join(self.app_dir, "LOG")
        os.makedirs(self.log_dir, exist_ok=True)
        self.system_log_path = os.path.join(
            self.log_dir,
            f"system_log{datetime.datetime.now():%Y%m%d}.txt"
        )
        
        # 初始化記憶體資料結構
        self.s1_isn_list = []
        self.s2_isn_list = []
        self.work_order_set = set()
        self.carton_scanned_count = 0
        self.load_label_data()
        
        # 建立 UI 畫面
        self.create_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        
    def load_configs(self):
        """讀取 config.ini 與 MySQLConfig.ini，並建立預設值防呆機制"""
        # --- 讀取 config.ini ---
        self.config = configparser.ConfigParser()
        config_file = os.path.join(self.app_dir, 'config.ini')
        if os.path.exists(config_file):
            self.config.read(config_file, encoding='utf-8')
        else:
            self.config['settings'] = {
                'csn_length': '18', 'mac_length': '12', 'Carton_ID': 'CAR260900001',
                'excel_name': '2544259S1S2.xlsx', 'label_name': 'Barcode label.lab',
                'json_name': 'label_data.json', 'log_file': 'scan_record.log',
                'tip_text': '請輸入或掃描資料（PANEL ID 23碼），並按 ENTER'
            }
            with open(config_file, 'w', encoding='utf-8') as f: self.config.write(f)

        # --- 讀取 MySQLConfig.ini ---
        self.mysql_config = configparser.ConfigParser()
        mysql_config_file = os.path.join(self.app_dir, 'MySQLConfig.ini')
        if os.path.exists(mysql_config_file):
            self.mysql_config.read(mysql_config_file, encoding='utf-8')
        else:
            self.mysql_config['setting'] = {
                'MySQL_FLAG': '1', 'MySQL_ServerIP': '10.4.5.13', 'MySQL_username': 'U94003',
                'MySQL_Password': 'U94003', 'MySQL_DB': 'LITEON', 'MySQL_PN': '2544259_S01+S2',
                'MySQL_LOT': '2544259-S01-003', 'MySQL_QTY': '80', 'MySQL_Carton': '1', 'MySQL_Carton_Serial': '1'
            }
            with open(mysql_config_file, 'w', encoding='utf-8') as f: self.mysql_config.write(f)

        # 擷取預設變數 (供畫面與邏輯使用)
        self.cfg_pn = self.mysql_config.get('setting', 'MySQL_PN', fallback='2544259_S01+S2')
        self.cfg_dc = datetime.datetime.now().strftime("%Y%m%d")
        self.cfg_lot = self.mysql_config.get('setting', 'MySQL_LOT', fallback='2544259-S01-003')
        self.cfg_qty = self.mysql_config.get('setting', 'MySQL_QTY', fallback='80')
        self.cfg_carton = self.mysql_config.get('setting', 'MySQL_Carton', fallback='1')
        self.cfg_c_serial = self.mysql_config.get('setting', 'MySQL_Carton_Serial', fallback='1')
        self.last_panel_id = self.config.get('settings', 'last_panel_id', fallback='')
        
    def create_widgets(self):
        """建立 GUI 介面，允許手動輸入修改"""
        padding = {'padx': 10, 'pady': 5}
        
        # 參數輸入區域
        frame_input = tk.LabelFrame(self.root, text="參數設定 (可修改)")
        frame_input.pack(fill="x", **padding)
        
        labels = ["P/N:", "D/C:", "LOT(WO#):", "Q'ty:", "MySQL_Carton:", "MySQL_Carton_Serial:"]
        self.entries = {}
        init_vals = [self.cfg_pn, self.cfg_dc, self.cfg_lot, self.cfg_qty, self.cfg_carton, self.cfg_c_serial]
        
        for i, text in enumerate(labels):
            row, col = divmod(i, 2)
            tk.Label(frame_input, text=text).grid(row=row, column=col*2, sticky="e", padx=5, pady=2)
            entry = tk.Entry(frame_input, width=20)
            entry.insert(0, init_vals[i])
            entry.grid(row=row, column=col*2+1, sticky="w", padx=5, pady=2)
            self.entries[text] = entry

        # Barcode 掃描區域
        frame_scan = tk.LabelFrame(self.root, text="條碼掃描輸入區")
        frame_scan.pack(fill="x", **padding)
        
        tk.Label(frame_scan, text="Panel ID / Barcode:").pack(side="left", padx=5, pady=5)
        self.entry_barcode = tk.Entry(frame_scan, width=40)
        self.entry_barcode.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.entry_barcode.insert(0, self.last_panel_id)
        self.entry_barcode.bind("<Return>", self.process_barcode)
        
        # 狀態與訊息顯示
        frame_status = tk.LabelFrame(self.root, text="系統提示與狀態")
        frame_status.pack(fill="both", expand=True, **padding)
        
        self.lbl_tip = tk.Label(frame_status, text=self.config.get('settings', 'tip_text', fallback='請掃描'), fg="blue", anchor="w")
        self.lbl_tip.pack(fill="x", padx=5, pady=2)
        
        self.txt_log = tk.Text(frame_status, height=10, state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=5, pady=5)
        self.root.after_idle(self.entry_barcode.focus_set)
        
    def write_system_log(self, message):
        """將訊息寫入當前目錄 LOG 資料夾中的 system_logyyyymmddhhmm.txt。"""
        try:
            with open(self.system_log_path, "a", encoding="utf-8") as log_file:
                log_file.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {message}\n")
        except OSError:
            pass

    def log_message(self, message, color=None):
        """將訊息列印到畫面的 Log 視窗中，並同步寫入 system log 檔案。"""
        self.txt_log.config(state="normal")
        log_text = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n"
        if color:
            tag_name = f"log_{color}"
            self.txt_log.tag_configure(tag_name, foreground=color)
            self.txt_log.insert(tk.END, log_text, tag_name)
        else:
            self.txt_log.insert(tk.END, log_text)
        self.txt_log.config(state="disabled")
        self.txt_log.see(tk.END)
        self.write_system_log(message)

    def log_exception(self, error):
        """將未預期錯誤寫入 system log 檔案，方便追查無主控台錯誤。"""
        self.write_system_log(f"{type(error).__name__}: {error}")

    def play_sound(self, sound_name):
        """播放掃描結果音效。"""
        sound_file = os.path.join(self.app_dir, sound_name)
        if not os.path.exists(sound_file):
            self.log_message(f"⚠️ 找不到音效檔案: {sound_file}")
            return
        try:
            winsound.PlaySound(sound_file, winsound.SND_FILENAME | winsound.SND_NODEFAULT)
        except (RuntimeError, OSError) as error:
            self.log_message(f"⚠️ 音效播放失敗: {error}")

    def process_barcode(self, event):
        """核心邏輯：當條碼按下 Enter 後觸發"""
        panel_id = self.entry_barcode.get().strip()
        self.entry_barcode.delete(0, tk.END)
        if not panel_id: return
        self.last_panel_id = panel_id
        self.config.set('settings', 'last_panel_id', panel_id)
        with open(os.path.join(self.app_dir, 'config.ini'), 'w', encoding='utf-8') as config_file:
            self.config.write(config_file)
        
        self.log_message(f"讀取到 Panel ID: {panel_id}")
        
        # 讀取畫面上最新修改的數值
        pn = self.entries["P/N:"].get().strip()
        lot = self.entries["LOT(WO#):"].get().strip()
        qty = int(self.entries["Q'ty:"].get().strip())
        carton = self.entries["MySQL_Carton:"].get().strip()
        c_serial = int(self.entries["MySQL_Carton_Serial:"].get().strip())
        if c_serial == 1:
            self.reset_label_progress()
        
        excel_name = self.config.get('settings', 'excel_name', fallback='2544259S1S2.xlsx').strip()
        excel_file = excel_name if os.path.isabs(excel_name) else os.path.join(self.app_dir, excel_name)
        
        # 2. 從 Excel 抓取對應的 iSN 
        if not os.path.exists(excel_file):
            messagebox.showerror("錯誤", f"找不到 Excel 檔案: {excel_name}\n搜尋路徑: {excel_file}")
            return
            
        try:
            from openpyxl import load_workbook
            from openpyxl.cell.cell import MergedCell
            workbook = load_workbook(excel_file)
            worksheet = workbook.active
            header_columns = {
                str(cell.value).strip(): cell.column
                for cell in worksheet[1]
                if cell.value is not None and str(cell.value).strip()
            }
            required_headers = ("工單號", "併板主PNL", "序號(S1SN)", "序號(S2SN)")
            missing_headers = [header for header in required_headers if header not in header_columns]
            if missing_headers:
                raise ValueError(f"Excel 缺少欄位: {', '.join(missing_headers)}")

            panel_column = header_columns["併板主PNL"]
            s1_column = header_columns["序號(S1SN)"]
            s2_column = header_columns["序號(S2SN)"]
            runtime_start = max(s1_column, s2_column) + 1
            runtime_headers = (
                "Flag", "刷Barcode時間", "箱號", "流水號",
                "SPI_T", "SPI_T_StopTime", "SPI_B", "SPI_B_StopTime",
                "AOI_T", "AOI_T_StopTime", "AOI_B", "AOI_B_StopTime",
                "AOI_B_B", "AOI_B_B_StopTime", "AOI_B_T", "AOI_B_T_StopTime",
            )
            runtime_columns = {
                header: runtime_start + offset
                for offset, header in enumerate(runtime_headers)
            }
            for header, column in runtime_columns.items():
                header_cell = worksheet.cell(row=1, column=column)
                if not isinstance(header_cell, MergedCell) and header_cell.value is None:
                    header_cell.value = header

            rows = [
                (row_number, worksheet[row_number])
                for row_number in range(2, worksheet.max_row + 1)
                if str(worksheet.cell(row=row_number, column=panel_column).value or '').strip() == panel_id
            ]
            if not rows:
                self.log_message(f"❌ Excel 中找不到該 Panel ID [{panel_id}] 的資料", "red")
                self.play_sound("buzz.wav")
                return

            successful_rows = []
            scan_failed = False
            skipped_rows = 0
            for row_number, row in rows:
                serials = [
                    (str(row[s1_column - 1].value or '').strip(), "S1"),
                    (str(row[s2_column - 1].value or '').strip(), "S2"),
                ]
                serials = [(isn, side) for isn, side in serials if isn]
                flag = str(row[runtime_columns["Flag"] - 1].value or '').strip()
                
                if flag == 'Y':
                    self.log_message(f"⚠️ PNL [{panel_id}] 已刷過主條碼 (Flag=Y)，跳過")
                    skipped_rows += 1
                    continue
                
                if not serials:
                    self.log_message(f"❌ PNL [{panel_id}] 沒有可驗證的 S1SN/S2SN，拒絕過站", "red")
                    scan_failed = True
                    break

                station_data_by_isn = {}
                for isn, side in serials:
                    station_data = self.verify_database(isn)
                    if not station_data:
                        self.log_message(f"❌ {side} iSN [{isn}] 未通過資料庫(SPI/AOI)檢驗，拒絕過站", "red")
                        scan_failed = True
                        break
                    station_data_by_isn[isn] = station_data
                if scan_failed:
                    break

                successful_rows.append((row_number, row, serials, station_data_by_isn))

            if not successful_rows and skipped_rows == len(rows):
                self.play_sound("buzz.wav")
                self.log_message(f"❌ Panel ID [{panel_id}] 已全部刷過，拒絕重複掃描", "red")
                return

            if scan_failed:
                self.play_sound("buzz.wav")
                self.log_message("❌ Panel ID 對應的 iSN 未全部過站成功，Excel 不回存", "red")
                return

            # 全部 iSN 都驗證成功後，才一次回存 Excel 與更新記憶體資料。
            for row_number, row, serials, station_data_by_isn in successful_rows:
                scan_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                worksheet.cell(row=row_number, column=runtime_columns["Flag"]).value = 'Y'
                worksheet.cell(row=row_number, column=runtime_columns["刷Barcode時間"]).value = scan_time
                worksheet.cell(row=row_number, column=runtime_columns["箱號"]).value = carton
                worksheet.cell(row=row_number, column=runtime_columns["流水號"]).value = c_serial
                for isn, side in serials:
                    if side == "S1":
                        if isn not in self.s1_isn_list: self.s1_isn_list.append(isn)
                    else:
                        if isn not in self.s2_isn_list: self.s2_isn_list.append(isn)
                    self.log_message(f"✅ {side} iSN [{isn}] 成功過站。流水號: {c_serial}")
                    c_serial += 1
                station_data = station_data_by_isn[serials[0][0]]
                for station_name in ("SPI_T", "SPI_B", "AOI_T", "AOI_B", "AOI_B_B", "AOI_B_T"):
                    worksheet.cell(row=row_number, column=runtime_columns[station_name]).value = station_data[station_name]
                    stop_time = station_data[f"{station_name}_StopTime"]
                    worksheet.cell(row=row_number, column=runtime_columns[f"{station_name}_StopTime"]).value = stop_time.strftime("%Y-%m-%d %H:%M:%S")

                row_wo = str(row[0].value or '').strip()
                self.work_order_set.add(row_wo)

            self.play_sound("pass.wav")
                
            # 寫回 Excel 檔案
            workbook.save(excel_file)
            self.entries["MySQL_Carton_Serial:"].delete(0, tk.END)
            self.entries["MySQL_Carton_Serial:"].insert(0, str(c_serial)) # 畫面流水號自動遞增
            self.save_runtime_settings()
            
            # 5 & 6. 寫入 json 資料
            self.write_label_json(pn, lot, qty)
            
            # 7. 檢查數量是否相等，觸發列印
            scanned_isn_count = sum(len(serials) for _, _, serials, _ in successful_rows)
            self.carton_scanned_count += scanned_isn_count
            self.log_message(f"本箱累計數量: {self.carton_scanned_count}/{qty}")
            if self.carton_scanned_count >= qty:
                self.trigger_print()
                
        except Exception as e:
            self.log_exception(e)
            messagebox.showerror("系統錯誤", f"處理過程中發生錯誤: {str(e)}")

    def save_runtime_settings(self):
        """保存目前畫面上的數量、箱號與流水號，供下次啟動載入。"""
        for entry_name, config_name in (
            ("Q'ty:", "MySQL_QTY"),
            ("MySQL_Carton:", "MySQL_Carton"),
            ("MySQL_Carton_Serial:", "MySQL_Carton_Serial"),
        ):
            self.mysql_config.set("setting", config_name, self.entries[entry_name].get().strip())
        with open(os.path.join(self.app_dir, "MySQLConfig.ini"), "w", encoding="utf-8") as config_file:
            self.mysql_config.write(config_file)
        self.config.set('settings', 'last_panel_id', self.last_panel_id)
        with open(os.path.join(self.app_dir, "config.ini"), "w", encoding="utf-8") as config_file:
            self.config.write(config_file)

    def close_application(self):
        self.save_runtime_settings()
        self.root.destroy()

    def load_label_data(self):
        """讀取上次保存的 QRCode 紀錄，讓重開後可繼續累加。"""
        json_file = os.path.join(
            self.app_dir,
            self.config.get('settings', 'json_name', fallback='label_data.json')
        )
        if not os.path.exists(json_file):
            return
        try:
            with open(json_file, 'r', encoding='utf-8') as data_file:
                data = json.load(data_file)
            self.work_order_set = set(data.get('WorkOrders', []))
            if data.get('S1QRCode1') is not None or data.get('S2QRCode1') is not None:
                self.s1_isn_list = [
                    isn for key in ('S1QRCode1', 'S1QRCode2', 'S1QRCode3', 'S1QRCode4')
                    for isn in str(data.get(key, '')).split(',') if isn.strip()
                ]
                self.s2_isn_list = [
                    isn for key in ('S2QRCode1', 'S2QRCode2', 'S2QRCode3', 'S2QRCode4')
                    for isn in str(data.get(key, '')).split(',') if isn.strip()
                ]
            elif len(self.work_order_set) == 1:
                self.s1_isn_list = [
                    isn for key in ('QRCode1', 'QRCode2', 'QRCode3', 'QRCode4')
                    for isn in str(data.get(key, '')).split(',') if isn.strip()
                ]
                self.s2_isn_list = []
            else:
                self.s1_isn_list = [
                    isn for key in ('QRCode1', 'QRCode3')
                    for isn in str(data.get(key, '')).split(',') if isn.strip()
                ]
                self.s2_isn_list = [
                    isn for key in ('QRCode2', 'QRCode4')
                    for isn in str(data.get(key, '')).split(',') if isn.strip()
                ]
            self.carton_scanned_count = len(self.s1_isn_list) + len(self.s2_isn_list)
        except (OSError, json.JSONDecodeError) as error:
            print(f"無法讀取 label_data.json: {error}")

    def verify_database(self, isn):
        """以 iSN 驗證 CARD 各站狀態，並回傳各站最新 StopTime。"""
        try:
            pymysql = importlib.import_module("pymysql")
            # 讀取資料庫連線參數
            conn = pymysql.connect(
                host=self.mysql_config.get('setting', 'MySQL_ServerIP'),
                user=self.mysql_config.get('setting', 'MySQL_username'),
                password=self.mysql_config.get('setting', 'MySQL_Password'),
                database=self.mysql_config.get('setting', 'MySQL_DB'),
                charset='utf8',
                cursorclass=pymysql.cursors.DictCursor
            )
            cursor = conn.cursor()
            
            query = """
                                SELECT c.SPI_T, c.SPI_B, c.AOI_T, c.AOI_B, c.AOI_B_B, c.AOI_B_T,
                                             (SELECT errorCode FROM SPI_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS SPI_T_ErrorCode,
                                             (SELECT errorCode FROM SPI_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS SPI_B_ErrorCode,
                                             (SELECT errorCode FROM AOI_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_T_ErrorCode,
                                             (SELECT errorCode FROM AOI_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_ErrorCode,
                                             (SELECT errorCode FROM AOI_B_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_B_ErrorCode,
                                             (SELECT errorCode FROM AOI_B_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_T_ErrorCode,
                                             (SELECT StopTime FROM SPI_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS SPI_T_StopTime,
                                             (SELECT StopTime FROM SPI_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS SPI_B_StopTime,
                                             (SELECT StopTime FROM AOI_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_T_StopTime,
                                             (SELECT StopTime FROM AOI_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_StopTime,
                                             (SELECT StopTime FROM AOI_B_B WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_B_StopTime,
                                             (SELECT StopTime FROM AOI_B_T WHERE iSN = c.iSN ORDER BY StopTime DESC, Serial DESC LIMIT 1) AS AOI_B_T_StopTime
                                FROM CARD c
                                WHERE c.iSN = %s
                LIMIT 1
            """
            cursor.execute(query, (isn,))
            result = cursor.fetchone()
            
            cursor.close()
            conn.close()
            station_names = ("SPI_T", "SPI_B", "AOI_T", "AOI_B", "AOI_B_B", "AOI_B_T")
            if not result:
                self.log_message(f"❌ iSN [{isn}] 未找到 CARD 資料或尚未寫入站別資料", "red")
                return None

            normalized_codes = {
                station: str(result.get(f"{station}_ErrorCode") or "").strip().upper()
                for station in station_names
            }

            failed_stations = []
            for station in station_names:
                card_value = result.get(station)
                code = normalized_codes[station]
                if card_value is None or card_value <= 100:
                    failed_stations.append(station)
                elif code not in ("PASS", "RPASS"):
                    failed_stations.append(station)

            if not failed_stations:
                return result

            failure_summary = "全部未通過" if len(failed_stations) == len(station_names) else ", ".join(failed_stations)
            self.log_message(
                f"❌ iSN [{isn}] 未通過資料庫(SPI/AOI)檢驗，未通過站別: {failure_summary}",
                "red"
            )
            return None
        except Exception as e:
            self.log_message(f"⚠️ 資料庫連線失敗或查詢錯誤: {e} (測試環境模擬通過)")
            return None

    def write_label_json(self, pn, lot, qty):
        """將 S1 與 S2 的 iSN 分組，格式化成 40 筆為一個 QRCode 欄位並存入 json"""
        json_file = os.path.join(
            self.app_dir,
            self.config.get('settings', 'json_name', fallback='label_data.json')
        )
        
        has_both_sides = bool(self.s1_isn_list and self.s2_isn_list)
        # 單一工單且只有一側時維持舊格式；S1/S2 成對資料固定分組。
        if len(self.work_order_set) == 1 and not has_both_sides:
            isn_list = self.s1_isn_list + self.s2_isn_list
            qr1_str = ",".join(isn_list[:40])
            qr2_str = ",".join(isn_list[40:80])
            qr3_str = ",".join(isn_list[80:120])
            qr4_str = ",".join(isn_list[120:160])
        else:
            qr1_str = ",".join(self.s1_isn_list[:40])
            qr2_str = ",".join(self.s2_isn_list[:40])
            qr3_str = ",".join(self.s1_isn_list[40:80])
            qr4_str = ",".join(self.s2_isn_list[40:80])

        data = {
            "PN": pn,
            "DateCode": self.entries["D/C:"].get().strip(),
            "LOT WO": lot,
            "QTY": str(qty),
            "QRCode1": qr1_str,
            "QRCode2": qr2_str,
            "QRCode3": qr3_str,
            "QRCode4": qr4_str,
            "WorkOrders": sorted(self.work_order_set)
            }
        if has_both_sides:
            data.update({
                "S1QRCode1": ",".join(self.s1_isn_list[:40]),
                "S1QRCode2": ",".join(self.s1_isn_list[40:80]),
                "S1QRCode3": ",".join(self.s1_isn_list[80:120]),
                "S1QRCode4": ",".join(self.s1_isn_list[120:160]),
                "S2QRCode1": ",".join(self.s2_isn_list[:40]),
                "S2QRCode2": ",".join(self.s2_isn_list[40:80]),
                "S2QRCode3": ",".join(self.s2_isn_list[80:120]),
                "S2QRCode4": ",".join(self.s2_isn_list[120:160]),
            })
        
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
    def trigger_print(self):
        """執行列印並確認已成功送入印表機佇列。"""
        self.log_message("達到目標數量 Q'ty，開始執行 test.bat 列印程序...")
        project_dir = self.app_dir
        batch_file = os.path.join(project_dir, "test.bat")
        output_file = os.path.join(project_dir, "temp_dos.txt")
        if os.path.exists(batch_file):
            try:
                if os.path.exists(output_file):
                    os.remove(output_file)
                result = subprocess.run(
                    ["cmd.exe", "/c", batch_file],
                    cwd=project_dir,
                    timeout=120,
                    check=False
                )
                output = ""
                if os.path.exists(output_file):
                    with open(output_file, "rb") as output_handle:
                        output = output_handle.read().decode("utf-8", errors="replace")

                if result.returncode != 0 or "PASS" not in output.upper() or "[OK]" not in output.upper():
                    self.play_sound("buzz.wav")
                    self.log_message("❌ FAIL：未確認列印已成功傳送至印表機佇列，維持目前箱號與流水號", "red")
                    return

                self.log_message("🚀 PASS [OK]：已成功傳送至印表機佇列。")
                carton = int(self.entries["MySQL_Carton:"].get().strip()) + 1
                self.entries["MySQL_Carton:"].delete(0, tk.END)
                self.entries["MySQL_Carton:"].insert(0, str(carton))
                if carton >= 2:
                    self.entries["MySQL_Carton_Serial:"].delete(0, tk.END)
                    self.entries["MySQL_Carton_Serial:"].insert(0, "1")
                self.save_runtime_settings()
                # 列印完清空計數
                self.s1_isn_list.clear()
                self.s2_isn_list.clear()
                self.work_order_set.clear()
                self.carton_scanned_count = 0
                self.clear_label_data()
            except subprocess.TimeoutExpired:
                self.play_sound("buzz.wav")
                self.log_message("❌ FAIL：列印程序逾時，維持目前箱號與流水號", "red")
            except Exception as e:
                self.log_message(f"❌ 列印執行失敗: {e}")
        else:
            self.log_message("❌ 找不到 test.bat 檔案，無法列印")

    def clear_label_data(self):
        """列印成功換箱後清除上一箱的 QRCode 暫存資料。"""
        json_file = os.path.join(
            self.app_dir,
            self.config.get('settings', 'json_name', fallback='label_data.json')
        )
        if not os.path.exists(json_file):
            return
        try:
            with open(json_file, 'r', encoding='utf-8') as data_file:
                data = json.load(data_file)
            for key in (
                'QRCode1', 'QRCode2', 'QRCode3', 'QRCode4',
                'S1QRCode1', 'S1QRCode2', 'S1QRCode3', 'S1QRCode4',
                'S2QRCode1', 'S2QRCode2', 'S2QRCode3', 'S2QRCode4',
            ):
                data[key] = ''
            data['WorkOrders'] = []
            with open(json_file, 'w', encoding='utf-8') as data_file:
                json.dump(data, data_file, ensure_ascii=False, indent=2)
        except (OSError, json.JSONDecodeError) as error:
            self.log_message(f"⚠️ 清除列印暫存資料失敗: {error}")

    def reset_label_progress(self):
        """流水號手動設為 1 時，清除目前箱內累計資料。"""
        self.s1_isn_list.clear()
        self.s2_isn_list.clear()
        self.work_order_set.clear()
        self.carton_scanned_count = 0
        self.clear_label_data()
        self.log_message("流水號為 1，已清空目前箱內計數，從 1 開始")


if __name__ == "__main__":
    root = tk.Tk()
    app = CartonSystemApp(root)
    root.mainloop()

