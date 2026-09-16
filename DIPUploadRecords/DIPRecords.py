import os
import configparser
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# 設定日誌存放的特定路徑
LOG_DIR = os.path.join(os.getcwd(), "LOG")
SYSTEM_LOG_DIR = os.path.join(os.getcwd(), "systemlog")
INI_FILE = os.path.join(os.getcwd(), "MySQLConfig.ini")

class BarcodeScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PANEL_ID 條碼掃描日誌生成器 (產線防重刷穩定版)")
        self.root.geometry("650x640")
        
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

        self.config = configparser.ConfigParser()
        self.config.optionxform = str # 保持 ini 欄位大小寫
        
        # 用於記錄這一次開機/未清空前，所有已經掃描過的不重複 panel_no
        self.scanned_panels = set()
        
        self.entries = {}
        self.create_widgets()
        self.load_ini_settings()
        
    def create_widgets(self):
        lbl_title = tk.Label(self.root, text="產線條碼掃描系統 (多版序防重刷模式)", font=("Microsoft JhengHei", 16, "bold"), fg="blue")
        lbl_title.pack(pady=10)
        
        form_frame = tk.Frame(self.root)
        form_frame.pack(pady=5, padx=20, fill="x")
        
        # 1. 機型 (TableDetailStr)
        frame1 = tk.Frame(form_frame)
        frame1.pack(fill="x", pady=4)
        tk.Label(frame1, text="機型 (TableDetailStr):", font=("Microsoft JhengHei", 10), width=30, anchor="w").pack(side="left")
        self.entries["machine_type"] = tk.Entry(frame1, font=("Microsoft JhengHei", 10), bd=2, relief="groove")
        self.entries["machine_type"].pack(side="right", fill="x", expand=True)
        self.entries["machine_type"].bind("<KeyRelease>", lambda e: self.update_preview())
        
        # 2. 工單 (MySQL_Job)
        frame2 = tk.Frame(form_frame)
        frame2.pack(fill="x", pady=4)
        tk.Label(frame2, text="工單 (MySQL_Job):", font=("Microsoft JhengHei", 10), width=30, anchor="w").pack(side="left")
        self.entries["work_order"] = tk.Entry(frame2, font=("Microsoft JhengHei", 10), bd=2, relief="groove")
        self.entries["work_order"].pack(side="right", fill="x", expand=True)
        self.entries["work_order"].bind("<KeyRelease>", lambda e: self.update_preview())
        
        # 3. 機種名稱面別 (MySQL_ModelName)
        frame3 = tk.Frame(form_frame)
        frame3.pack(fill="x", pady=4)
        tk.Label(frame3, text="機種名稱面別 (MySQL_ModelName):", font=("Microsoft JhengHei", 10), width=30, anchor="w").pack(side="left")
        self.entries["model_name"] = tk.Entry(frame3, font=("Microsoft JhengHei", 10), bd=2, relief="groove")
        self.entries["model_name"].pack(side="right", fill="x", expand=True)
        self.entries["model_name"].bind("<KeyRelease>", lambda e: self.update_preview())

        # 4. 版序 (MySQL_Preface)
        frame4 = tk.Frame(form_frame)
        frame4.pack(fill="x", pady=4)
        tk.Label(frame4, text="版序數量 (MySQL_Preface):", font=("Microsoft JhengHei", 10), width=30, anchor="w").pack(side="left")
        self.entries["version_count"] = tk.Entry(frame4, font=("Microsoft JhengHei", 10), bd=2, relief="groove")
        self.entries["version_count"].pack(side="right", fill="x", expand=True)
        self.entries["version_count"].bind("<KeyRelease>", lambda e: self.update_preview())
        
        # 分隔線
        lbl_line = tk.Label(form_frame, text="─"*50, fg="gray")
        lbl_line.pack(fill="x", pady=8)
        
        # 5. PANELNO (主掃描核心)
        frame5 = tk.Frame(form_frame)
        frame5.pack(fill="x", pady=4)
        tk.Label(frame5, text="★ 掃描 PANELNO (持續輸入):", font=("Microsoft JhengHei", 11, "bold"), fg="red", width=30, anchor="w").pack(side="left")
        self.entries["panel_no"] = tk.Entry(frame5, font=("Microsoft JhengHei", 12, "bold"), bd=3, relief="sunken", bg="#FFFFCC")
        self.entries["panel_no"].pack(side="right", fill="x", expand=True)
        self.entries["panel_no"].bind("<Return>", self.process_panel_scan)
        self.entries["panel_no"].bind("<KeyRelease>", lambda e: self.update_preview())
        
        # 預覽區域
        preview_frame = tk.LabelFrame(self.root, text=" 目前設定資料與檔名生成預覽 ", font=("Microsoft JhengHei", 10, "bold"), fg="green")
        preview_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(preview_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.txt_preview = tk.Text(preview_frame, font=("Consolas", 10), bg="#f5f5f5", state="disabled", yscrollcommand=scrollbar.set)
        self.txt_preview.pack(fill="both", expand=True, padx=5, pady=5)
        scrollbar.config(command=self.txt_preview.yview)
        
        # 清除按鈕
        btn_clear = tk.Button(self.root, text="清除並重置防重刷", font=("Microsoft JhengHei", 10), bg="#f44336", fg="white", width=18, command=self.clear_all)
        btn_clear.pack(pady=5)

    def load_ini_settings(self):
        """相容讀取：先嘗試 utf-8，失敗則切換至 cp950 (ANSI)"""
        if os.path.exists(INI_FILE):
            encodings = ['utf-8', 'cp950', 'gbk', 'utf-8-sig']
            success = False
            for enc in encodings:
                try:
                    self.config.read(INI_FILE, encoding=enc)
                    success = True
                    break
                except Exception:
                    continue
            
            if success and 'setting' in self.config:
                sec = self.config['setting']
                self.entries["machine_type"].insert(0, sec.get("TableDetailStr", ""))
                self.entries["work_order"].insert(0, sec.get("MySQL_Job", ""))
                self.entries["model_name"].insert(0, sec.get("MySQL_ModelName", ""))
                self.entries["version_count"].insert(0, sec.get("MySQL_Preface", "1"))
        
        self.entries["panel_no"].focus_set()
        self.update_preview()

    def save_ini_settings(self, machine_type, work_order, model_name, version_count):
        """寫入 INI：使用 cp950 (ANSI) 儲存以最大化相容原舊有系統"""
        try:
            if not os.path.exists(INI_FILE):
                self.config['setting'] = {}
            else:
                encodings = ['utf-8', 'cp950', 'gbk']
                for enc in encodings:
                    try:
                        self.config.read(INI_FILE, encoding=enc)
                        break
                    except Exception:
                        continue
                        
                if 'setting' not in self.config:
                    self.config['setting'] = {}
                    
            self.config['setting']['TableDetailStr'] = machine_type
            self.config['setting']['MySQL_Job'] = work_order
            self.config['setting']['MySQL_ModelName'] = model_name
            self.config['setting']['MySQL_Preface'] = version_count
            
            with open(INI_FILE, 'w', encoding='cp950') as configfile:
                self.config.write(configfile)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法更新 MySQLConfig.ini，原因:\n{e}")

    def update_preview(self, override_panel_no=None):
        """即時顯示預覽與檔名，支援傳入剛刷完的真實 PANELNO"""
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        
        machine_type = self.entries["machine_type"].get().strip()
        work_order = self.entries["work_order"].get().strip()
        model_name = self.entries["model_name"].get().strip()
        version_str = self.entries["version_count"].get().strip()
        
        # 如果有傳入剛剛成功生產的 panel_no 就強制鎖定顯示它，否則從輸入框抓取
        if override_panel_no is not None:
            panel_no = override_panel_no
        else:
            panel_no = self.entries["panel_no"].get().strip()
            
        status = "PASS"
        
        if model_name and not model_name.endswith('B'):
            model_name += 'B'
            
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        generated_files = [] # 修正處：將變數明確提至 try 的外部宣告，保障全域安全性
        
        try:
            v_count = int(version_str) if version_str else 1
            if v_count < 1: v_count = 1
        except ValueError:
            v_count = 1
        
        # 顯示目前的防重刷計數
        self.txt_preview.insert(tk.END, f"【防重刷計數】 當前已累計刷入: {len(self.scanned_panels)} 片不重複面板\n")
        self.txt_preview.insert(tk.END, f"【設定狀態】 狀態固定: {status} | 連噴總檔數 (版序): {v_count} 個檔案\n")
        self.txt_preview.insert(tk.END, f"【設定欄位】 機型: {machine_type} | 工單: {work_order} | 機種: {model_name}\n")
        self.txt_preview.insert(tk.END, f"【面板條碼】 {panel_no if panel_no else '(等待掃描...)'}\n")
        self.txt_preview.insert(tk.END, f"【儲存路徑】 {LOG_DIR}\n")
        self.txt_preview.insert(tk.END, f"─"*50 + "\n")
        
        if override_panel_no is not None:
            self.txt_preview.insert(tk.END, f"✅【剛才成功產出的完整檔名清單（共 {v_count} 個檔案）】:\n")
        else:
            self.txt_preview.insert(tk.END, f"【預計產出檔名清單（共 {v_count} 個檔案）】:\n")
            
        for idx in range(1, v_count + 1):
            display_panel = panel_no if panel_no else "[請掃描PANELNO]"
            file_components = [machine_type, work_order, display_panel, status, str(idx), model_name, current_time]
            filename = "_".join([c for c in file_components if c]) + ".txt"
            self.txt_preview.insert(tk.END, f" 📄 {filename}\n")
            
        self.txt_preview.config(state="disabled")
        return machine_type, work_order, model_name, v_count, panel_no, status

    def log_to_system_log(self, generated_filenames):
        """將產出的檔名附加記錄到 \systemlog\以今天日期命名.log 中"""
        try:
            today_str = datetime.now().strftime("%Y%m%d") # 例如 20260916
            sys_log_file = os.path.join(SYSTEM_LOG_DIR, f"{today_str}.log")
            time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            with open(sys_log_file, "a", encoding="utf-8") as sf:
                for fname in generated_filenames:
                    sf.write(f"[{time_now}] 產生檔案: {fname}\n")
        except Exception as e:
            print(f"寫入系統日誌系統失敗: {e}")
            
    def process_panel_scan(self, event):
        """當 PANELNO 欄位掃描到 Enter 時觸發"""
        # 生產前先確實提取目前輸入框的真實文字
        generated_files = []
        success_count = 0
        machine_type, work_order, model_name, v_count, panel_no, status = self.update_preview()
        
        if not machine_type or not work_order or not model_name:
            messagebox.showwarning("警告", "請先確認『機型』、『工單』、『機種名稱面別』皆已輸入！")
            return
        if not panel_no:
            return
            
        # 核心防重刷邏輯
        if panel_no in self.scanned_panels:
            self.root.bell() 
            messagebox.showerror("重複掃描警告", f"❌ 條碼重複！\nPANELNO: {panel_no}\n此面板在此班別已掃描並產生過 LOG 檔案！")
            self.entries["panel_no"].delete(0, tk.END)
            self.update_preview()
            return

        self.save_ini_settings(machine_type, work_order, model_name, str(v_count))
        
        success_count = 0
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        
        try:
            for idx in range(1, v_count + 1):
                file_components = [machine_type, work_order, panel_no, status, str(idx), model_name, current_time]
                filename = "_".join(file_components) + ".txt"
                file_path = os.path.join(LOG_DIR, filename)
                
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(f"建立時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"PANEL_NO: {panel_no}\n")
                    f.write(f"STATUS: {status}\n")
                    f.write(f"PREFACE_INDEX: {idx}/{v_count}\n")
                
                generated_files.append(filename)
                success_count += 1
                
            # 將產出的檔案名寫入系統紀錄檔 \systemlog\YYYYMMDD.log
            self.log_to_system_log(generated_files)
            self.scanned_panels.add(panel_no)    
            self.entries["panel_no"].delete(0, tk.END)
            self.update_preview()
            
            self.txt_preview.config(state="normal")
            self.txt_preview.insert("1.0", f"✨ [成功連噴] 一鍵成功產生 {success_count} 個 Log 檔案！ (時間戳: {current_time})\n\n")
            self.txt_preview.config(state="disabled")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"無法完全寫入 Log 檔案，原因:\n{e}")
    def clear_all(self):
        """清空所有輸入框"""
        for entry in self.entries.values():
            entry.delete(0, tk.END)
            self.update_preview()
            self.entries["machine_type"].focus_set()
				
if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeScannerApp(root)
    root.mainloop()