import streamlit as st
import openpyxl
import json
import os
import re
from datetime import datetime, date
import calendar
import io

# ================= 系統設定 =================
DATA_FILE = "MONSTER 隊員場地費.xlsx"  
HISTORY_FILE = "attendance_history.json" 

def get_role_from_fill(fill):
    """根據 Excel 儲存格顏色判斷身分"""
    if not fill or not fill.start_color:
        return "未分類"
    c = fill.start_color
    if c.type == 'theme':
        if c.theme == 4:
            return "🟦 底層"
        elif c.theme == 5:
            return "🟥 上層"
    return "未分類"

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_history(history):
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=4)

def main():
    st.set_page_config(page_title="MONSTER 點名系統", page_icon="🏐", layout="wide")
    
    # ================= 手機版介面優化 =================
    st.markdown("""
        <style>
        .stMultiSelect span { font-size: 18px !important; }
        .stButton>button {
            width: 100%; height: 50px; font-size: 18px;
            font-weight: bold; border-radius: 8px;
        }
        </style>
    """, unsafe_allow_html=True)
    # =================================================

    st.title("🏐 MONSTER 線上點名系統")
    
    # 顯示成功訊息 (獨立於重整機制之外)
    if "success_msg" in st.session_state:
        st.success(st.session_state["success_msg"])
        del st.session_state["success_msg"]

    # 1. 載入 Excel 檔案
    if not os.path.exists(DATA_FILE):
        st.error(f"❌ 找不到檔案 `{DATA_FILE}`，請確認檔案已放入資料夾！")
        return
        
    try:
        wb = openpyxl.load_workbook(DATA_FILE)
    except Exception as e:
        st.error(f"❌ 檔案讀取失敗！如果您剛才遇到錯誤，請先使用備份檔覆蓋壞掉的 Excel。詳細錯誤: {e}")
        return

    # 2. 選擇日期與確認工作表
    date_selected = st.date_input("📅 選擇點名日期", date.today())
    target_sheet_name = date_selected.strftime("%Y%m")
    
    sheet_exists = target_sheet_name in wb.sheetnames
    sheet_for_read = wb[target_sheet_name] if sheet_exists else wb[wb.sheetnames[-1]]

    # 3. 解析名單與當日點名狀況
    members = []
    attendees_today = []
    stats = {"🟦 底層": 0, "🟥 上層": 0, "👤 外賓": 0}
    date_col = None
    
    # 尋找今日日期欄位
    if sheet_exists:
        for c in range(3, 34):
            c_val = sheet_for_read.cell(row=2, column=c).value
            if isinstance(c_val, datetime) and c_val.date() == date_selected:
                date_col = c
                break

    # 讀取名單
    for r in range(4, sheet_for_read.max_row + 1):
        val = sheet_for_read.cell(row=r, column=2).value
        if not val: continue
        name = str(val).strip()
        if not name: continue
        
        fill = sheet_for_read.cell(row=r, column=2).fill
        role = get_role_from_fill(fill)
        
        if role != "未分類": 
            members.append({"name": name, "role": role})
            
        # 統計今日出席
        if date_col and sheet_for_read.cell(row=r, column=date_col).value == 1:
            display_role = "👤 外賓" if role == "未分類" else role
            stats[display_role] += 1
            attendees_today.append({"name": name, "role": display_role, "row": r})

    # 歷史排序
    history = load_history()
    bases = [m["name"] for m in members if m["role"] == "🟦 底層"]
    flyers = [m["name"] for m in members if m["role"] == "🟥 上層"]
    bases.sort(key=lambda x: history.get(x, 0), reverse=True)
    flyers.sort(key=lambda x: history.get(x, 0), reverse=True)

    st.divider()

    # ================= 雙欄設計 =================
    col_left, col_right = st.columns([1.2, 1])

    # ---------- 【左側】新增點名區 ----------
    with col_left:
        st.header("📝 新增點名")
        with st.form("attendance_form"):
            selected_bases = st.multiselect("🟦 選擇底層人員", options=bases)
            selected_flyers = st.multiselect("🟥 選擇上層人員", options=flyers)
            
            st.markdown("---")
            guests_input = st.text_area("👤 填寫臨打/外賓 (請用逗號或換行分隔)")
            guests_count = st.number_input("未具名外賓人數 (自動產生代號)", min_value=0, value=0)
            
            # 【優化】直接將執行邏輯綁定在按鈕判斷式內
            if st.form_submit_button("✅ 送出並寫入 Excel"):
                guests = [g.strip() for g in re.split(r'[,\s\n]+', guests_input) if g.strip()]
                if not guests and guests_count > 0:
                    guests = [f"外賓_{i+1}" for i in range(guests_count)]
                    
                all_attendees = selected_bases + selected_flyers
                
                if not all_attendees and not guests:
                    st.warning("⚠️ 您尚未選擇或輸入任何出席人員喔！")
                else:
                    try:
                        # ======== 自動換月功能 ========
                        if not sheet_exists:
                            source_sheet = wb[wb.sheetnames[-1]]
                            new_sheet = wb.copy_worksheet(source_sheet)
                            new_sheet.title = target_sheet_name
                            days_in_month = calendar.monthrange(date_selected.year, date_selected.month)[1]
                            for c in range(3, 34): 
                                day = c - 2
                                if day <= days_in_month:
                                    new_sheet.cell(row=2, column=c, value=datetime(date_selected.year, date_selected.month, day))
                                else:
                                    new_sheet.cell(row=2, column=c, value=None)
                                for r in range(4, new_sheet.max_row + 1):
                                    new_sheet.cell(row=r, column=c, value=None)
                            target_sheet = new_sheet
                        else:
                            target_sheet = wb[target_sheet_name]

                        # 找尋欄位並寫入
                        target_col = None
                        for c in range(3, 34):
                            c_val = target_sheet.cell(row=2, column=c).value
                            if isinstance(c_val, datetime) and c_val.date() == date_selected:
                                target_col = c
                                break
                                
                        if not target_col:
                            st.error(f"❌ 找不到 {date_selected} 對應的日期欄位！")
                        else:
                            name_to_row = {}
                            for r in range(4, target_sheet.max_row + 1):
                                val = target_sheet.cell(row=r, column=2).value
                                if val: name_to_row[str(val).strip()] = r
                                
                            for name in all_attendees:
                                if name in name_to_row:
                                    target_sheet.cell(row=name_to_row[name], column=target_col).value = 1
                                    
                            for guest in guests:
                                new_r = target_sheet.max_row + 1
                                target_sheet.cell(row=new_r, column=1).value = "x"
                                target_sheet.cell(row=new_r, column=2).value = guest
                                target_sheet.cell(row=new_r, column=target_col).value = 1
                                
                            wb.save(DATA_FILE)
                            for p in (all_attendees + guests):
                                history[p] = history.get(p, 0) + 1
                            save_history(history)
                            
                            st.session_state["success_msg"] = f"🎉 成功新增 {len(all_attendees) + len(guests)} 人的點名紀錄！"
                            st.rerun()
                    except PermissionError:
                        st.error("❌ 無法寫入！請檢查您是否正在電腦上打開著 Excel 檔案？")
                    except Exception as e:
                        st.error(f"❌ 發生錯誤: {e}")

    # ---------- 【右側】當日狀況與「批次刪除」 ----------
    with col_right:
        st.header("📊 當日已點名名單")
        
        st.markdown(f"### 總人數：{len(attendees_today)} 人")
        
        bases_present = [p['name'] for p in attendees_today if p['role'] == "🟦 底層"]
        flyers_present = [p['name'] for p in attendees_today if p['role'] == "🟥 上層"]
        guests_present = [p['name'] for p in attendees_today if p['role'] == "👤 外賓"]
        
        st.markdown(f"🟦 **底層 ({stats['🟦 底層']})：** " + "、".join(bases_present))
        st.markdown(f"🟥 **上層 ({stats['🟥 上層']})：** " + "、".join(flyers_present))
        if guests_present:
            st.markdown(f"👤 **外賓 ({stats['👤 外賓']})：** " + "、".join(guests_present))
        
        if not attendees_today:
            st.info("今日尚無人點名。")
        else:
            st.write("---")
            with st.form("delete_form"):
                st.markdown("⚠️ **發現點錯人？在這裡取消點名**")
                # 加上專屬 key 確保狀態不遺失
                to_delete = st.multiselect(
                    "選擇要「刪除/取消」的人員 (可一次選多個)", 
                    options=[p['name'] for p in attendees_today],
                    key="delete_select"
                )
                
                # 【優化】直接將刪除邏輯綁定在按鈕內
                if st.form_submit_button("❌ 刪除所選人員並更新"):
                    if not to_delete:
                        st.warning("⚠️ 請先選擇要刪除的人員！")
                    else:
                        try:
                            for name in to_delete:
                                r = next(p['row'] for p in attendees_today if p['name'] == name)
                                # 徹底清除 Excel 中的出席標記 (設定為空)
                                sheet_for_read.cell(row=r, column=date_col).value = None
                            
                            wb.save(DATA_FILE)
                            st.session_state["success_msg"] = f"🗑️ 已成功刪除 {len(to_delete)} 人的出席紀錄！"
                            st.rerun()
                        except PermissionError:
                            st.error("❌ 無法寫入！請檢查您是否正在電腦上打開著 Excel 檔案？")
                        except Exception as e:
                            st.error(f"❌ 刪除時發生錯誤: {e}")

    # 備份下載
    st.divider()
    with open(DATA_FILE, "rb") as f:
        st.download_button(
            label="📥 下載最新的 Excel 檔案 (備份/查閱用)",
            data=f,
            file_name=f"{date_selected}_{DATA_FILE}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
