import configparser
from datetime import datetime
import pathlib
import re
import sys
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd
import pymysql
from openpyxl import load_workbook


BASE_DIR = pathlib.Path(__file__).resolve().parent
TEMPLATE_PATTERN = "*shipping list.xlsx"

DB_COLUMNS = [
    "COMPLETE_TIME", "RELATION_ORDER", "DN_ITEM", "CUST_PART_NO", "PART_NO",
    "CUSTOMER_SN", "TEMP_SN", "CUSTOMER_SN2", "REEL_NO", "QTY", "PALLET_NO",
    "CARTON_NO", "STATUS", "WORK_ORDER", "MAC", "MODEL_NAME", "MFG_Date",
    "Shipping_Date", "PO",
]

# Template headings may use either of these names.
ALIASES = {
    "MODELNAME": "MODEL_NAME",
    "MODEL_NAME": "MODEL_NAME",
    "SHIPPINGDATE": "Shipping_Date",
    "MFGDATE": "MFG_Date",
}


def read_db_config(config_file="MySQLConfig.ini"):
    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding="utf-8")
        return {
            "host": config.get("setting", "MySQL_ServerIP"),
            "user": config.get("setting", "MySQL_username"),
            "password": config.get("setting", "MySQL_Password"),
            "database": config.get("setting", "MySQL_DB"),
        }
    except Exception as exc:
        messagebox.showerror("配置錯誤", f"無法讀取設定檔 {config_file}\n原因: {exc}")
        raise


def increment_pallet_no(base_str, increment_by):
    if not base_str:
        return ""
    match = re.match(r"(.*?)(\d+)$", str(base_str))
    if match:
        prefix, number = match.groups()
        return f"{prefix}{int(number) + increment_by:0{len(number)}d}"
    return f"{base_str}_{increment_by}" if increment_by else str(base_str)


def normalise_heading(value):
    if value is None:
        return ""
    text = re.sub(r"[^a-zA-Z0-9]", "", str(value)).upper()
    return ALIASES.get(text, str(value).strip())


def find_template():
    files = sorted(BASE_DIR.glob(TEMPLATE_PATTERN))
    if not files:
        files = sorted(pathlib.Path.cwd().glob(TEMPLATE_PATTERN))
    return files[0] if files else None


def read_template_definition():
    """Return template headings, v-selected headings, and ModelName choices.

    The shipping-list workbook is intentionally treated as a template: a cell
    containing v marks the column that should be exported.  The heading is
    taken from the same row or the closest non-empty cell above it.  This also
    works with the common layout where headings are on row 2 and data starts on
    row 3.
    """
    template = find_template()
    if not template:
        return DB_COLUMNS[:], DB_COLUMNS[:], []

    workbook = load_workbook(template, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30), values_only=True))
    heading_row = None
    model_col = None
    for row_number, row in enumerate(rows, 1):
        for col_number, value in enumerate(row, 1):
            if normalise_heading(value) == "MODEL_NAME":
                heading_row, model_col = row_number, col_number
                break
        if heading_row:
            break

    if heading_row is None:
        workbook.close()
        return DB_COLUMNS[:], DB_COLUMNS[:], []

    headings = {}
    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(heading_row, col).value
        if value not in (None, ""):
            headings[col] = normalise_heading(value)

    selected = []
    # Usually v is in the row immediately below the headings. Search the
    # header area so minor template row shifts do not break the export.
    for col, heading in headings.items():
        marked = any(str(sheet.cell(row, col).value).strip().lower() == "v"
                     for row in range(1, min(heading_row + 3, sheet.max_row + 1)))
        if marked and heading in DB_COLUMNS and heading not in selected:
            selected.append(heading)

    if not selected:
        selected = [heading for heading in headings.values() if heading in DB_COLUMNS]
    if not selected:
        selected = DB_COLUMNS[:]

    choices = []
    for row in range(max(heading_row + 1, 3), sheet.max_row + 1):
        value = sheet.cell(row, model_col).value
        if value not in (None, "") and str(value).strip().lower() != "v":
            text = str(value).strip()
            if text not in choices:
                choices.append(text)
    workbook.close()
    return selected, selected, choices


def execute_export():
    work_order = entry_wo.get().strip()
    model_name = combo_model.get().strip()
    qty_text = entry_qty.get().strip()
    input_po = entry_po.get().strip()
    input_complete_time = entry_complete_time.get().strip()
    input_relation_order = entry_relation_order.get().strip()
    input_dn_item = entry_cs_shipping_notice.get().strip()
    input_part_no = entry_product_part_no.get().strip()
    input_shipping_date = entry_shipping_date.get().strip()
    start_pallet = entry_pallet.get().strip()
    cartons_text = entry_carton_count.get().strip()

    if not work_order:
        messagebox.showwarning("警告", "請輸入工單號碼！")
        return
    try:
        requested_qty = int(qty_text)
        if requested_qty <= 0:
            raise ValueError
    except ValueError:
        messagebox.showwarning("輸入錯誤", "QTY 必須為大於 0 的整數！")
        return

    export_columns, _, _ = read_template_definition()
    db_config = read_db_config()
    select_fields = ", ".join(DB_COLUMNS)
    sql = f"""SELECT {select_fields} FROM ASN_Log
              WHERE WORK_ORDER = %s
                AND CARTON_NO IS NOT NULL
                AND TRIM(CARTON_NO) <> ''"""
    params = [work_order]
    if model_name:
        sql += " AND MODEL_NAME = %s"
        params.append(model_name)
    sql += " ORDER BY CARTON_NO ASC"

    connection = None
    try:
        label_status.config(text="狀態：正在查詢資料庫...", fg="blue")
        root.update_idletasks()
        connection = pymysql.connect(**db_config, charset="utf8")
        df = pd.read_sql(sql, connection, params=tuple(params))
        if df.empty:
            label_status.config(text="狀態：查無資料", fg="orange")
            messagebox.showinfo("提示", "找不到符合條件且 CARTON_NO 有值的資料。")
            return

        # QTY is explicitly supplied by the operator and is written to every
        # generated shipping-list row, rather than trusting the DB value.
        df["QTY"] = requested_qty
        overrides = {
            "PO": input_po, "COMPLETE_TIME": input_complete_time,
            "RELATION_ORDER": input_relation_order, "DN_ITEM": input_dn_item,
            "PART_NO": input_part_no, "Shipping_Date": input_shipping_date,
        }
        for column, value in overrides.items():
            if value:
                df[column] = value

        if start_pallet and cartons_text:
            try:
                per_pallet = int(cartons_text)
                if per_pallet <= 0:
                    raise ValueError
                df["PALLET_NO"] = [increment_pallet_no(start_pallet, i // per_pallet)
                                    for i in range(len(df))]
            except ValueError:
                messagebox.showwarning("輸入錯誤", "幾個 CARTON 必須為大於 0 的整數，已保留資料庫棧板號。")

        output = df[[column for column in export_columns if column in df.columns]].copy()
        file_name = f"{work_order}_{datetime.now():%Y%m%d}.xlsx"
        output.to_excel(file_name, index=False, engine="openpyxl")
        label_status.config(text="狀態：匯出成功！", fg="green")
        messagebox.showinfo("成功", f"檔案匯出成功！\n檔名：{file_name}\n共 {len(output)} 筆資料。")
    except pymysql.MySQLError as exc:
        label_status.config(text="狀態：資料庫錯誤", fg="red")
        messagebox.showerror("資料庫錯誤", f"操作失敗：\n{exc}")
    except Exception as exc:
        label_status.config(text="狀態：發生錯誤", fg="red")
        messagebox.showerror("錯誤", f"發生預期之外的錯誤：\n{exc}")
    finally:
        if connection:
            connection.close()


def focus_next(widget):
    widget.focus_set()


selected_columns, _, model_choices = read_template_definition()
root = tk.Tk()
root.title("ASN Log 資料匯出工具")
root.geometry("520x650")
root.resizable(False, False)
frame = ttk.Frame(root, padding="20")
frame.pack(fill=tk.BOTH, expand=True)

tk.Label(frame, text="工單資料 Excel 匯出系統", font=("Microsoft JhengHei", 14, "bold")).pack(pady=(0, 10))


def input_row(label, default="", widget_type="entry", values=()):
    row = ttk.Frame(frame)
    row.pack(fill=tk.X, pady=3)
    tk.Label(row, text=label, font=("Microsoft JhengHei", 10), width=18, anchor="w").pack(side=tk.LEFT)
    if widget_type == "combo":
        widget = ttk.Combobox(row, values=values, font=("Microsoft JhengHei", 10))
    else:
        widget = ttk.Entry(row, font=("Microsoft JhengHei", 10))
    if default:
        widget.insert(0, default)
    widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
    return widget


entry_wo = input_row("請輸入工單號碼：", "WOTQ7553D")
entry_po = input_row("請輸入 PO 號碼：")
combo_model = input_row("ModelName：", widget_type="combo", values=model_choices)
entry_qty = input_row("QTY（產生數量）：", "1")
entry_complete_time = input_row("COMPLETE_TIME：")
entry_relation_order = input_row("RELATION_ORDER：")
entry_cs_shipping_notice = input_row("CS出貨通知：")
entry_product_part_no = input_row("成品料號：")
entry_shipping_date = input_row("Shipping_Date：")
entry_pallet = input_row("PALLET_NO 起始值：", "PL001")
entry_carton_count = input_row("幾個 CARTON 換棧板：", "2")

label_status = tk.Label(frame, text="狀態：準備就緒", font=("Microsoft JhengHei", 9), fg="gray")
label_status.pack(pady=5)
ttk.Button(frame, text="開始查詢並匯出 Excel", command=execute_export).pack(fill=tk.X, ipady=5)

entries = [entry_wo, entry_po, combo_model, entry_qty, entry_complete_time,
           entry_relation_order, entry_cs_shipping_notice, entry_product_part_no,
           entry_shipping_date, entry_pallet, entry_carton_count]
for current, next_widget in zip(entries, entries[1:]):
    current.bind("<Return>", lambda event, widget=next_widget: focus_next(widget))
entries[-1].bind("<Return>", lambda event: execute_export())
entry_wo.focus_set()
root.mainloop()
