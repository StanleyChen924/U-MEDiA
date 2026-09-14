import os
import datetime

def log_and_print(message, log_file_path):
    """同時將訊息輸出到控制台並寫入 LOG 檔案"""
    print(message)
    try:
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception as e:
        print(f"❌ 無法寫入 LOG 檔案: {e}")

def rename_spi_by_sequence(aoi_dir, spi_dir):
    """
    依據【一對一順序】將 Source_LOG 的 PANELNO 替換到 Modify_LOG 的檔名中。
    若數量不一致，不中斷程式，而是比對並撈出不匹配的 PANELNO 差異。
    """
    # 1. 初始化 LOG 資料夾與檔案名稱
    log_dir = "LOG"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    current_date = datetime.datetime.now().strftime("%Y%m%d")
    log_file_name = f"{current_date}.log"
    log_file_path = os.path.join(log_dir, log_file_name)

    # 寫入開始標頭
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_and_print(f"\n==========================================", log_file_path)
    log_and_print(f"🚀 程式啟動時間: {current_time}", log_file_path)
    log_and_print(f"==========================================", log_file_path)

    if not os.path.exists(aoi_dir) or not os.path.exists(spi_dir):
        log_and_print(f"❌ 錯誤：Source_LOG 或 Modify_LOG 資料夾路徑不存在。", log_file_path)
        return

    # 2. 讀取並【排序】Source 檔案，解析出 Source 的 PANELNO
    aoi_files = sorted([f for f in os.listdir(aoi_dir) if os.path.isfile(os.path.join(aoi_dir, f))])
    source_panels = []
    for f in aoi_files:
        parts = os.path.splitext(f)[0].split('_')
        source_panels.append(parts[2] if len(parts) >= 3 else None)

    # 3. 讀取並【排序】Modify 檔案，解析出 Modify 的 PANELNO
    spi_files = sorted([f for f in os.listdir(spi_dir) if os.path.isfile(os.path.join(spi_dir, f))])
    modify_panels = []
    for f in spi_files:
        parts = os.path.splitext(f)[0].split('_')
        modify_panels.append(parts[2] if len(parts) >= 3 else None)

    log_and_print(f"\n📊 檔案數量檢查：Source 共有 {len(aoi_files)} 個 | Modify 共有 {len(spi_files)} 個", log_file_path)
    
    # 🚨 💡 核心抓漏邏輯：比對兩邊的 PANELNO 集合差異
    if len(aoi_files) != len(spi_files):
        log_and_print(f"\n⚠️ 【警告】兩邊檔案數量不一致！開始撈出異常的 PANELNO...", log_file_path)
        
        # 轉成集合(Set)進行差集運算 (排除 None 的無效資料)
        set_source = set([p for p in source_panels if p])
        set_modify = set([p for p in modify_panels if p])
        
        # 狀況 A：Source 有，但 Modify 沒有 (代表 Modify 漏存了或不見了)
        missing_in_modify = set_source - set_modify
        if missing_in_modify:
            log_and_print(f"🔍 檢查結果：Modify_LOG 中【搞丟/遺漏】了以下 PANELNO 的檔案：", log_file_path)
            for p in sorted(list(missing_in_modify)):
                log_and_print(f"   ➡️ 缺失 PANELNO: {p}", log_file_path)
        
        # 狀況 B：Modify 有，但 Source 沒有 (可能 AOI 漏檢測或漏存)
        missing_in_source = set_modify - set_source
        if missing_in_source:
            log_and_print(f"🔍 檢查結果：Source_LOG 中【缺失】了以下 PANELNO 的檔案：", log_file_path)
            for p in sorted(list(missing_in_source)):
                log_and_print(f"   ➡️ 缺失 PANELNO: {p}", log_file_path)
        
        log_and_print(f"\n💡 提示：因數量不一致，接下來將僅針對「前 {min(len(aoi_files), len(spi_files))} 個檔案」依序強制配對修改！\n", log_file_path)

    # 4. 依據順序一對一替換檔名
    match_count = min(len(source_panels), len(spi_files))
    if match_count == 0:
        log_and_print("ℹ️ 無可配對的檔案，程式結束。", log_file_path)
        return

    log_and_print("===== 開始依順序替換 Modify_LOG 檔名 =====", log_file_path)
    success_count = 0
    
    for i in range(match_count):
        current_panel_no = source_panels[i]
        spi_filename = spi_files[i]
        
        if not current_panel_no:
            log_and_print(f"ℹ️ 順序 {i+1:03d}: 對應的 Source 檔名解析失敗，跳過 Modify 檔案 [{spi_filename}]", log_file_path)
            continue
            
        spi_name_without_ext, spi_ext = os.path.splitext(spi_filename)
        spi_parts = spi_name_without_ext.split('_')
        
        if len(spi_parts) >= 3:
            spi_parts[2] = current_panel_no  
            new_spi_filename = "_".join(spi_parts) + spi_ext
            
            old_file_full_path = os.path.join(spi_dir, spi_filename)
            new_file_full_path = os.path.join(spi_dir, new_spi_filename)
            
            try:
                os.rename(old_file_full_path, new_file_full_path)
                log_and_print(f"✨ [順序 {i+1:03d}] 成功更名: [{spi_filename}] ➡️ [{new_spi_filename}]", log_file_path)
                success_count += 1
            except Exception as e:
                log_and_print(f"❌ [順序 {i+1:03d}] 檔案 [{spi_filename}] 更名失敗: {e}", log_file_path)
        else:
            log_and_print(f"⚠️ 順序 {i+1:03d}: Modify 檔名格式結構不符，無法替換: [{spi_filename}]", log_file_path)

    log_and_print(f"\n===== 處理完成！共依序成功修改 {success_count} 個 Modify 檔案 =====", log_file_path)
    log_and_print(f"💾 完整日誌已儲存至: {log_file_path}\n", log_file_path)

if __name__ == "__main__":
    rename_spi_by_sequence(r"Source_LOG", r"Modify_LOG")
