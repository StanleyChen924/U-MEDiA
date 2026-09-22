import configparser
from datetime import datetime
import re
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import pymysql


def read_db_config(config_file="MySQLConfig.ini"):
    """讀取 INI 設定檔"""
    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding="utf-8")
        db_info = {
            "host": config.get("setting", "MySQL_ServerIP"),
            "user": config.get("setting", "MySQL_username"),
            "password": config.get("setting", "MySQL_Password"),
            "database": config.get("setting", "MySQL_DB"),
        }
        return db_info
    except Exception as e:
        messagebox.showerror("配置錯誤", f"無法讀取設定檔 {config_file}\n原因: {e}")
        sys.exit(1)


def increment_pallet_no(base_str, increment_by):
    """將文字型態或數字型態的棧板號加上指定的增量"""
    if not base_str:
        return ""

    match = re.match(r"(.*?)(\d+)$", str(base_str))
    if match:
        prefix, num_str = match.groups()
        length = len(num_str)
        new_num = int(num_str) + increment_by
        return f"{prefix}{new_num:0{length}d}"
    else:
        return f"{base_str}_{increment_by}" if increment_by > 0 else base_str


def execute_export():
    """按鈕觸發事件"""
    work_order = entry_wo.get().strip()
    input_po = entry_po.get().strip()
    input_complete_time = entry_complete_time.get().strip()
    input_relation_order = entry_relation_order.get().strip()
    input_cs_shipping_notice = entry_cs_shipping_notice.get().strip()
    input_product_part_no = entry_product_part_no.get().strip()
    input_shipping_date = entry_shipping_date.get().strip()
    start_pallet = entry_pallet.get().strip()
    cartons_per_pallet_str = entry_carton_count.get().strip()

    if not work_order:
        messagebox.showwarning("警告", "請輸入工單號碼！")
        return

    db_config = read_db_config()

    columns = [
        "COMPLETE_TIME",
        "RELATION_ORDER",
        "DN_ITEM",
        "CUST_PART_NO",
        "PART_NO",
        "CUSTOMER_SN",
        "TEMP_SN",
        "CUSTOMER_SN2",
        "REEL_NO",
        "QTY",
        "PALLET_NO",
        "CARTON_NO",
        "STATUS",
        "WORK_ORDER",
        "MAC",
        "MODEL_NAME",
        "MFG_Date",
        "Shipping_Date",
        "PO",
    ]

    select_fields = ", ".join(columns)
    sql = f"""
        SELECT {select_fields} 
        FROM ASN_Log 
        WHERE WORK_ORDER = %s 
        ORDER BY CARTON_NO ASC
    """

    connection = None
    try:
        label_status.config(text="狀態：正在查詢資料庫...", fg="blue")
        root.update()

        connection = pymysql.connect(
            host=db_config["host"],
            user=db_config["user"],
            password=db_config["password"],
            database=db_config["database"],
            charset="utf8",
        )

        df = pd.read_sql(sql, connection, params=(work_order,))

        if df.empty:
            label_status.config(text="狀態：查無資料", fg="orange")
            messagebox.showinfo("提示", f"找不到符合工單號碼【{work_order}】的資料。")
            return

        if input_po:
            df["PO"] = input_po

        input_fields = {
            "COMPLETE_TIME": input_complete_time,
            "RELATION_ORDER": input_relation_order,
            "DN_ITEM": input_cs_shipping_notice,
            "PART_NO": input_product_part_no,
            "Shipping_Date": input_shipping_date,
        }
        for column, value in input_fields.items():
            if value:
                df[column] = value

        if start_pallet and cartons_per_pallet_str:
            try:
                cartons_per_pallet = int(cartons_per_pallet_str)
                if cartons_per_pallet <= 0:
                    raise ValueError

                pallet_list = []
                for i in range(len(df)):
                    increment_times = i // cartons_per_pallet
                    current_pallet = increment_pallet_no(start_pallet, increment_times)
                    pallet_list.append(current_pallet)

                df["PALLET_NO"] = pallet_list

            except ValueError:
                messagebox.showwarning(
                    "輸入錯誤", "「幾個 CARTON」必須為大於 0 的整數，將不自動計算棧板號。"
                )

        today_str = datetime.now().strftime("%Y%m%d")
        file_name = f"{work_order}_{today_str}.xlsx"

        df.to_excel(file_name, index=False, columns=columns, engine="openpyxl")

        label_status.config(text="狀態：匯出成功！", fg="green")
        messagebox.showinfo(
            "成功", f"檔案匯出成功！\n檔名：{file_name}\n共 {len(df)} 筆資料。"
        )
        """
        for entry in (
            entry_wo,
            entry_po,
            entry_complete_time,
            entry_relation_order,
            entry_cs_shipping_notice,
            entry_product_part_no,
            entry_shipping_date,
            entry_pallet,
            entry_carton_count,
        ):
            entry.delete(0, tk.END)
            """
        entry_wo.focus_set()

    except pymysql.MySQLError as e:
        label_status.config(text="狀態：資料庫錯誤", fg="red")
        messagebox.showerror("資料庫錯誤", f"操作失敗：\n{e}")
    except Exception as e:
        label_status.config(text="狀態：發生錯誤", fg="red")
        messagebox.showerror("錯誤", f"發生預期之外的錯誤：\n{e}")
    finally:
        if connection:
            connection.close()


def focus_next_widget(current_entry):
    """將焦點切換到下一個指定的輸入框"""
    current_entry.focus()


# ==================== UI 視窗介面設計 ====================
root = tk.Tk()
root.title("ASN Log 資料匯出工具 (流暢輸入版)")
root.geometry("520x570")
root.resizable(False, False)

# 視窗置中
window_width = 480
window_height = 570
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)
root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

frame = ttk.Frame(root, padding="20")
frame.pack(fill=tk.BOTH, expand=True)

# 標題
label_title = tk.Label(
    frame, text="工單資料 Excel 匯出系統", font=("Microsoft JhengHei", 14, "bold")
)
label_title.pack(pady=(0, 10))

# --- 工單輸入 ---
wo_frame = ttk.Frame(frame)
wo_frame.pack(fill=tk.X, pady=5)
tk.Label(
    wo_frame, text="請輸入工單號碼：", font=("Microsoft JhengHei", 10), width=18, anchor="w"
).pack(side=tk.LEFT)
entry_wo = ttk.Entry(wo_frame, font=("Microsoft JhengHei", 10))
entry_wo.insert(0, "WOTQ7553D")
entry_wo.pack(side=tk.LEFT, fill=tk.X, expand=True)
entry_wo.focus()

# --- PO 輸入 ---
po_frame = ttk.Frame(frame)
po_frame.pack(fill=tk.X, pady=5)
tk.Label(
    po_frame, text="請輸入 PO 號碼：", font=("Microsoft JhengHei", 10), width=18, anchor="w"
).pack(side=tk.LEFT)
entry_po = ttk.Entry(po_frame, font=("Microsoft JhengHei", 10))
entry_po.pack(side=tk.LEFT, fill=tk.X, expand=True)


def create_input_row(label_text):
    row_frame = ttk.Frame(frame)
    row_frame.pack(fill=tk.X, pady=3)
    tk.Label(
        row_frame,
        text=label_text,
        font=("Microsoft JhengHei", 10),
        width=18,
        anchor="w",
    ).pack(side=tk.LEFT)
    entry = ttk.Entry(row_frame, font=("Microsoft JhengHei", 10))
    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    return entry


# --- ASN 欄位輸入 ---
entry_complete_time = create_input_row("COMPLETE_TIME：")
entry_relation_order = create_input_row("RELATION_ORDER：")
entry_cs_shipping_notice = create_input_row("CS出貨通知：")
entry_product_part_no = create_input_row("成品料號：")
entry_shipping_date = create_input_row("Shipping_Date：")

# --- 棧板起始值輸入 ---
pallet_frame = ttk.Frame(frame)
pallet_frame.pack(fill=tk.X, pady=5)
tk.Label(
    pallet_frame,
    text="PALLET_NO 起始值：",
    font=("Microsoft JhengHei", 10),
    width=18,
    anchor="w",
).pack(side=tk.LEFT)
entry_pallet = ttk.Entry(pallet_frame, font=("Microsoft JhengHei", 10))
entry_pallet.insert(0, "PL001")
entry_pallet.pack(side=tk.LEFT, fill=tk.X, expand=True)

# --- 幾箱換棧板輸入 ---
carton_frame = ttk.Frame(frame)
carton_frame.pack(fill=tk.X, pady=5)
tk.Label(
    carton_frame,
    text="幾個 CARTON 換棧板：",
    font=("Microsoft JhengHei", 10),
    width=18,
    anchor="w",
).pack(side=tk.LEFT)
entry_carton_count = ttk.Entry(carton_frame, font=("Microsoft JhengHei", 10))
entry_carton_count.insert(0, "2")
entry_carton_count.pack(side=tk.LEFT, fill=tk.X, expand=True)

# --- 【核心新增：Enter 鍵切換焦點與執行】 ---
# 工單欄位按 Enter -> 跳到 PO 欄位
entry_wo.bind("<Return>", lambda event: focus_next_widget(entry_po))
# PO 欄位按 Enter -> 跳到 COMPLETE_TIME 欄位
entry_po.bind("<Return>", lambda event: focus_next_widget(entry_complete_time))
# ASN 欄位按 Enter 依序切換焦點
entry_complete_time.bind("<Return>", lambda event: focus_next_widget(entry_relation_order))
entry_relation_order.bind("<Return>", lambda event: focus_next_widget(entry_cs_shipping_notice))
entry_cs_shipping_notice.bind("<Return>", lambda event: focus_next_widget(entry_product_part_no))
entry_product_part_no.bind("<Return>", lambda event: focus_next_widget(entry_shipping_date))
entry_shipping_date.bind("<Return>", lambda event: focus_next_widget(entry_pallet))
# 棧板起始值按 Enter -> 跳到 幾個 CARTON
entry_pallet.bind("<Return>", lambda event: focus_next_widget(entry_carton_count))
# 最後一個欄位按 Enter -> 直接執行匯出
entry_carton_count.bind("<Return>", lambda event: execute_export())


# 狀態顯示
label_status = tk.Label(frame, text="狀態：準備就緒", font=("Microsoft JhengHei", 9), fg="gray")
label_status.pack(pady=5)

# 執行按鈕
btn_export = ttk.Button(frame, text="開始查詢並匯出 Excel", command=execute_export)
btn_export.pack(fill=tk.X, ipady=5)

root.mainloop()
