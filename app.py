import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime, date
import calendar

# ================= 系統設定 =================
# ⚠️ 請確保這裡是你最新的 Google Sheet 網址！
SHEET_URL = "https://docs.google.com/spreadsheets/d/1GuaV0Rwdj3MYHQXL_HNtE3ge0s4eaOEVkpT9TbsKYtI/edit?gid=1805401804#gid=1805401804"

# ================= 連線到 Google Sheets =================
@st.cache_resource
def init_connection():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    skey = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(skey, scopes=scopes)
    gc = gspread.authorize(credentials)
    return gc

def get_workbook():
    gc = init_connection()
    return gc.open_by_url(SHEET_URL)

def main():
    st.set_page_config(page_title="MONSTER 點名系統", page_icon="🏐", layout="wide")
    
    st.markdown("""
        <style>
        .stMultiSelect span { font-size: 18px !important; }
        .stButton>button { width: 100%; height: 50px; font-size: 18px; font-weight: bold; border-radius: 8px; }
        </style>
    """, unsafe_allow_html=True)

    # ================= 1. 登入系統 =================
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        st.title("🔒 MONSTER 點名系統 - 登入")
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("請輸入幹部帳號與密碼以進入系統")
            user_id = st.selectbox("👤 選擇點名帳號", ["A", "B", "C", "D", "E"])
            password = st.text_input("🔑 密碼", type="password")
            
            if st.button("✅ 登入系統"):
                if password == "MONSTER":
                    st.session_state["logged_in"] = True
                    st.session_state["current_user"] = user_id
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤！請重新輸入。")
        return 

    # ================= 已登入狀態 =================
    st.title("🏐 MONSTER 線上點名系統 ☁️")
    col_title, col_logout = st.columns([8, 1])
    with col_title:
        st.caption(f"目前登入者：**{st.session_state['current_user']}** ｜ 資料庫：Google Sheets (性別極速版)")
    with col_logout:
        if st.button("登出"):
            st.session_state["logged_in"] = False
            st.rerun()
    
    if "success_msg" in st.session_state:
        st.success(st.session_state["success_msg"])
        del st.session_state["success_msg"]

    try:
        wb = get_workbook()
    except Exception as e:
        st.error(f"❌ 無法連線至 Google Sheets，請檢查 Secrets 金鑰是否正確。錯誤訊息: {e}")
        return

    # 隱藏資料庫 System_DB (現在只用來記錄「點名次數」作排序用，不再記錄身分)
    try:
        db_sheet = wb.worksheet("System_DB")
    except gspread.exceptions.WorksheetNotFound:
        db_sheet = wb.add_worksheet("System_DB", rows=200, cols=3)
        db_sheet.append_row(["姓名", "點名次數"])

    db_records = db_sheet.get_all_records()
    db_map = {str(r.get('姓名', '')): int(r.get('點名次數', 0)) for r in db_records if r.get('姓名')}

    tab_attendance, tab_manage = st.tabs(["📝 點名作業", "⚙️ 名單管理"])

    # ================== 分頁 1：點名作業 ==================
    with tab_attendance:
        date_selected = st.date_input("📅 選擇點名日期", date.today())
        target_sheet_name = date_selected.strftime("%Y%m")
        
        sheet_names = [s.title for s in wb.worksheets()]
        sheet_exists = target_sheet_name in sheet_names
        sheet_for_read = wb.worksheet(target_sheet_name) if sheet_exists else wb.worksheet(sheet_names[0] if "System" in sheet_names[-1] else sheet_names[-1])

        # 一次抓取整頁資料 (光速)
        all_values = sheet_for_read.get_all_values()
        
        members = []
        attendees_today = []
        claimed_tapes = []
        current_taker = "無"
        guests_today = 0
        purchased_tapes_today = 0
        
        # 尋找目標日期所在的欄位 (因為多了C欄性別，所以第一天 1號 從 D欄(index 3) 開始)
        target_col_idx = date_selected.day + 2 
        
        if sheet_exists and len(all_values) >= 2:
            try:
                current_taker = str(all_values[0][target_col_idx]).strip()
            except:
                pass
                
        tape_col_idx = None
        if len(all_values) >= 3:
            for i, val in enumerate(all_values[2]): # 第三列找「白貼」
                if str(val).strip() == "白貼":
                    tape_col_idx = i
                    break

        missing_in_db = []
        for r_idx, row in enumerate(all_values[3:], start=4): # 第四列開始是名單
            if len(row) < 2: continue
            name = str(row[1]).strip()
            if not name: continue
            
            # 外賓與白貼購買專屬列
            if name == "外賓人數":
                if sheet_exists and target_col_idx < len(row) and row[target_col_idx]:
                    try: guests_today = int(row[target_col_idx])
                    except: pass
                continue
            if name == "購買白貼":
                if sheet_exists and target_col_idx < len(row) and row[target_col_idx]:
                    try: purchased_tapes_today = int(row[target_col_idx])
                    except: pass
                continue

            # ✨ 全新邏輯：透過 C 欄(index 2) 的文字「男/女」來判斷身分
            gender = str(row[2]).strip() if len(row) > 2 else ""
            if gender == "男":
                role = "🟦 底層"
            elif gender == "女":
                role = "🟥 上層"
            else:
                role = "未分類"

            if name not in db_map:
                missing_in_db.append([name, 0])
                db_map[name] = 0
            
            count = db_map[name]
            members.append({"name": name, "role": role, "row": r_idx, "count": count})
            
            if sheet_exists and target_col_idx < len(row) and str(row[target_col_idx]).strip() == "1":
                attendees_today.append({"name": name, "role": role, "row": r_idx})
                
            if tape_col_idx and tape_col_idx < len(row) and str(row[tape_col_idx]).strip().upper() == "V":
                claimed_tapes.append(name)
                
        if missing_in_db:
            db_sheet.append_rows(missing_in_db)

        # 排序
        bases = [m["name"] for m in members if m["role"] == "🟦 底層"]
        flyers = [m["name"] for m in members if m["role"] == "🟥 上層"]
        bases.sort(key=lambda x: db_map.get(x, 0), reverse=True)
        flyers.sort(key=lambda x: db_map.get(x, 0), reverse=True)
        
        available_for_tape = [m['name'] for m in members if m['name'] not in claimed_tapes]
        available_for_tape.sort(key=lambda x: db_map.get(x, 0), reverse=True)

        st.divider()
        col_left, col_right = st.columns([1.2, 1])

        with col_left:
            st.header("📝 新增點名與白貼登記")
            with st.form("attendance_form"):
                selected_bases = st.multiselect("🟦 選擇底層人員", options=bases)
                selected_flyers = st.multiselect("🟥 選擇上層人員", options=flyers)
                
                st.markdown("---")
                col_g, col_t = st.columns(2)
                with col_g:
                    guests_input_count = st.number_input("👤 當日外賓總人數", min_value=0, value=guests_today)
                with col_t:
                    tapes_purchased_input = st.number_input("📦 當日購買白貼數量", min_value=0, value=purchased_tapes_today)
                
                st.markdown("---")
                selected_tapes = st.multiselect("🩹 登記領取白貼 (隱藏本月已領者)", options=available_for_tape)
                
                if st.form_submit_button("✅ 送出並同步 Google Sheets"):
                    all_attendees = selected_bases + selected_flyers
                    try:
                        if not sheet_exists:
                            valid_sheets = [s for s in sheet_names if re.match(r'^\d{6}$', s)]
                            valid_sheets.sort()
                            source_sheet = wb.worksheet(valid_sheets[-1]) if valid_sheets else sheet_for_read
                            
                            new_sheet = wb.duplicate_sheet(source_sheet.id, new_sheet_name=target_sheet_name)
                            days_in_month = calendar.monthrange(date_selected.year, date_selected.month)[1]
                            
                            # ✨ 換月清空邏輯：現在日期欄位是從 D欄 (4) 到 AH欄 (34)
                            cell_updates = []
                            for c_idx in range(4, 35): 
                                day = c_idx - 3
                                if day <= days_in_month:
                                    cell_updates.append(gspread.Cell(row=2, col=c_idx, value=f"{date_selected.month}/{day}"))
                                else:
                                    cell_updates.append(gspread.Cell(row=2, col=c_idx, value=""))
                                
                                cell_updates.append(gspread.Cell(row=1, col=c_idx, value="")) 
                                for r_idx in range(4, new_sheet.row_count + 1):
                                    cell_updates.append(gspread.Cell(row=r_idx, col=c_idx, value=""))
                                    
                            if tape_col_idx:
                                for r_idx in range(4, new_sheet.row_count + 1):
                                    cell_updates.append(gspread.Cell(row=r_idx, col=tape_col_idx + 1, value=""))
                                    
                            new_sheet.update_cells(cell_updates)
                            target_sheet = new_sheet
                        else:
                            target_sheet = wb.worksheet(target_sheet_name)

                        t_col = target_col_idx + 1 # Gspread API 是 1-based
                        updates = []
                        updates.append(gspread.Cell(row=1, col=t_col, value=st.session_state["current_user"]))
                        
                        name_to_row = {m['name']: m['row'] for m in members}
                            
                        for name in all_attendees:
                            if name in name_to_row:
                                updates.append(gspread.Cell(row=name_to_row[name], col=t_col, value=1))
                                
                        if tape_col_idx:
                            t_tape_col = tape_col_idx + 1
                            for name in selected_tapes:
                                if name in name_to_row:
                                    updates.append(gspread.Cell(row=name_to_row[name], col=t_tape_col, value="V"))
                        
                        # 處理外賓與購買
                        guest_r = next((i for i, row in enumerate(all_values) if len(row)>1 and str(row[1]).strip()=="外賓人數"), None)
                        if not guest_r:
                            guest_r = target_sheet.row_count + 1
                            target_sheet.update(values=[["x", "外賓人數"]], range_name=f"A{guest_r}:B{guest_r}")
                        else: guest_r += 1 
                        updates.append(gspread.Cell(row=guest_r, col=t_col, value=guests_input_count if guests_input_count > 0 else ""))
                        
                        purchase_r = next((i for i, row in enumerate(all_values) if len(row)>1 and str(row[1]).strip()=="購買白貼"), None)
                        if not purchase_r:
                            purchase_r = target_sheet.row_count + 1
                            target_sheet.update(values=[["x", "購買白貼"]], range_name=f"A{purchase_r}:B{purchase_r}")
                        else: purchase_r += 1
                        updates.append(gspread.Cell(row=purchase_r, col=t_col, value=tapes_purchased_input if tapes_purchased_input > 0 else ""))
                        
                        target_sheet.update_cells(updates)
                        
                        # 更新點名次數資料庫
                        db_all = db_sheet.get_all_records()
                        count_updates = []
                        for i, r in enumerate(db_all, start=2):
                            if r.get('姓名') in all_attendees:
                                count_updates.append(gspread.Cell(row=i, col=2, value=int(r.get('點名次數', 0)) + 1))
                        if count_updates:
                            db_sheet.update_cells(count_updates)
                        
                        msg = f"🎉 成功同步點名紀錄！負責人：{st.session_state['current_user']} "
                        if selected_tapes: msg += f" | 已登記 {len(selected_tapes)} 人領取白貼。"
                        st.session_state["success_msg"] = msg
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 發生錯誤: {e}")

        with col_right:
            st.header("📊 當日與本月概況")
            st.info(f"📝 本日點名負責人：**{current_taker}**")
            
            bases_present = [p['name'] for p in attendees_today if p['role'] == "🟦 底層"]
            flyers_present = [p['name'] for p in attendees_today if p['role'] == "🟥 上層"]
            
            st.markdown(f"**總出席隊員：{len(attendees_today)} 人**")
            st.markdown(f"🟦 **底層/男 ({len(bases_present)})：** " + "、".join(bases_present))
            st.markdown(f"🟥 **上層/女 ({len(flyers_present)})：** " + "、".join(flyers_present))
            
            col_s1, col_s2 = st.columns(2)
            col_s1.markdown(f"👤 **外賓人數：** `{guests_today}` 人")
            col_s2.markdown(f"📦 **購買白貼：** `{purchased_tapes_today}` 捲")
            
            st.write("---")
            st.markdown(f"🩹 **本月已領白貼人員 ({len(claimed_tapes)})**")
            if claimed_tapes:
                st.markdown("、".join(claimed_tapes))
            else:
                st.caption("本月尚無人領取")
            
            st.write("---")
            with st.form("delete_form"):
                st.markdown("⚠️ **發現登記錯誤？在這裡取消**")
                to_delete_attend = st.multiselect("取消「點名」", options=[p['name'] for p in attendees_today])
                to_delete_tape = st.multiselect("取消本月「白貼領取」", options=claimed_tapes)
                
                if st.form_submit_button("❌ 取消所選項並更新"):
                    if not to_delete_attend and not to_delete_tape:
                        st.warning("⚠️ 請先選擇要取消的項目！")
                    else:
                        try:
                            t_col = target_col_idx + 1
                            del_updates = []
                            for name in to_delete_attend:
                                r = next(m['row'] for m in members if m['name'] == name)
                                del_updates.append(gspread.Cell(row=r, col=t_col, value=""))
                            if tape_col_idx:
                                t_tape_col = tape_col_idx + 1
                                for name in to_delete_tape:
                                    r = next(m['row'] for m in members if m['name'] == name)
                                    del_updates.append(gspread.Cell(row=r, col=t_tape_col, value=""))
                                    
                            sheet_for_read.update_cells(del_updates)
                            st.session_state["success_msg"] = f"🗑️ 成功取消所選紀錄！"
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 取消時發生錯誤: {e}")

    # ================== 分頁 2：名單管理 ==================
    with tab_manage:
        st.header("⚙️ 隊員名單增刪管理")
        st.info("💡 系統已升級「男/女」極速判斷！新增與修改人員時，系統會自動幫您將性別填入 C 欄。")
        col_add, col_edit, col_del = st.columns(3)
        
        with col_add:
            with st.form("add_member_form"):
                st.subheader("➕ 新增人員")
                new_name = st.text_input("輸入新隊員姓名")
                new_role = st.selectbox("選擇身分", ["🟦 底層", "🟥 上層"])
                
                if st.form_submit_button("新增至 G-Sheets"):
                    new_name = new_name.strip()
                    if not new_name:
                        st.warning("請輸入姓名！")
                    elif new_name in [m['name'] for m in members]:
                        st.warning("此人員已存在！")
                    else:
                        try:
                            ref_row = 4
                            for m in reversed(members):
                                if m['role'] == new_role:
                                    ref_row = m['row']
                                    break
                            insert_idx = ref_row + 1
                            
                            body = {
                                "requests": [
                                    {"insertDimension": {"range": {"sheetId": sheet_for_read.id, "dimension": "ROWS", "startIndex": insert_idx - 1, "endIndex": insert_idx}}},
                                    {"copyPaste": {"source": {"sheetId": sheet_for_read.id, "startRowIndex": ref_row - 1, "endRowIndex": ref_row},
                                                   "destination": {"sheetId": sheet_for_read.id, "startRowIndex": insert_idx - 1, "endRowIndex": insert_idx},
                                                   "pasteType": "PASTE_NORMAL"}}
                                ]
                            }
                            wb.client.batch_update(wb.id, body)
                            
                            # ✨ 自動寫入男/女
                            gender_val = "男" if new_role == "🟦 底層" else "女"
                            sheet_for_read.update(values=[["x", new_name, gender_val]], range_name=f"A{insert_idx}:C{insert_idx}")
                            
                            # 清空後面的打卡紀錄 (從 D 欄開始清空)
                            blanks = [[""] * 31]
                            sheet_for_read.update(values=blanks, range_name=f"D{insert_idx}:AH{insert_idx}")
                            if tape_col_idx:
                                sheet_for_read.update_cell(insert_idx, tape_col_idx + 1, "")
                                
                            db_sheet.append_row([new_name, 0])
                            
                            st.session_state["success_msg"] = f"✅ 成功將 {new_name} 加入名單！"
                            st.rerun()
                        except Exception as e:
                            st.error(f"新增失敗：{e}")

        with col_edit:
            with st.form("edit_member_form"):
                st.subheader("✏️ 修改身分/姓名")
                old_name = st.selectbox("選擇要修改的隊員", sorted([m['name'] for m in members]))
                edit_new_name = st.text_input("輸入正確姓名 (若不改名請留空)")
                edit_role = st.selectbox("修改身分", ["🟦 底層", "🟥 上層"], index=0)
                
                if st.form_submit_button("儲存修改"):
                    try:
                        target_name = edit_new_name.strip() if edit_new_name.strip() else old_name
                        gender_val = "男" if edit_role == "🟦 底層" else "女"
                        r = next(m['row'] for m in members if m['name'] == old_name)
                        
                        # ✨ 一次性把新名字跟性別(男/女)寫入 B 欄與 C 欄
                        sheet_for_read.update(values=[[target_name, gender_val]], range_name=f"B{r}:C{r}")
                            
                        # 更新資料庫
                        cell = db_sheet.find(old_name, in_column=1)
                        if cell:
                            db_sheet.update_cell(cell.row, 1, target_name)
                        
                        st.session_state["success_msg"] = f"✏️ 成功更新「{target_name}」的資料！"
                        st.rerun()
                    except Exception as e:
                        st.error(f"修改失敗：{e}")

        with col_del:
            with st.form("delete_member_form"):
                st.subheader("🗑️ 刪除人員")
                st.error("刪除後將完全移除！")
                del_name = st.selectbox("選擇要永久刪除的隊員", sorted([m['name'] for m in members]))
                
                if st.form_submit_button("永久刪除該隊員"):
                    try:
                        r = next(m['row'] for m in members if m['name'] == del_name)
                        sheet_for_read.delete_rows(r)
                        
                        cell = db_sheet.find(del_name, in_column=1)
                        if cell: db_sheet.delete_rows(cell.row)
                        
                        st.session_state["success_msg"] = f"🗑️ 已將 {del_name} 徹底移除！"
                        st.rerun()
                    except Exception as e:
                        st.error(f"刪除失敗：{e}")

if __name__ == "__main__":
    main()
