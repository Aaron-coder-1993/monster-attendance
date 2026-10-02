import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import re
from datetime import datetime, date
import calendar

# ================= 系統設定 =================
# ⚠️ 請確保這裡是你最新的 Google Sheet 網址！
SHEET_URL = "https://docs.google.com/spreadsheets/d/1GuaV0Rwdj3MYHQXL_HNtE3ge0s4eaOEVkpT9TbsKYtI/edit?usp=drivesdk"
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
        st.caption(f"目前登入者：**{st.session_state['current_user']}** ｜ 資料庫：Google Sheets (綽號顯示版)")
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

    try:
        db_sheet = wb.worksheet("System_DB")
    except gspread.exceptions.WorksheetNotFound:
        db_sheet = wb.add_worksheet("System_DB", rows=200, cols=3)
        db_sheet.append_row(["姓名(綽號)", "點名次數"])

    db_records = db_sheet.get_all_records()
    db_map = {str(r.get('姓名(綽號)', r.get('姓名', ''))): int(r.get('點名次數', 0)) for r in db_records}

    tab_attendance, tab_manage = st.tabs(["📝 點名作業", "⚙️ 名單管理"])

    # ================== 分頁 1：點名作業 ==================
    with tab_attendance:
        date_selected = st.date_input("📅 選擇點名日期", date.today())
        target_sheet_name = date_selected.strftime("%Y%m")
        
        sheet_names = [s.title for s in wb.worksheets()]
        sheet_exists = target_sheet_name in sheet_names
        sheet_for_read = wb.worksheet(target_sheet_name) if sheet_exists else wb.worksheet(sheet_names[0] if "System" in sheet_names[-1] else sheet_names[-1])

        all_values = sheet_for_read.get_all_values()
        
        members = []
        attendees_today = []
        claimed_tapes = []
        current_taker = "無"
        guests_today = 0
        purchased_tapes_today = 0
        
        # ✨ 日期欄位偏移：因為 B=本名, C=綽號, D=性別，所以 E欄 (index 4) 才是 1 號
        target_col_idx = date_selected.day + 3 
        
        if sheet_exists and len(all_values) >= 2:
            try:
                current_taker = str(all_values[0][target_col_idx]).strip()
            except:
                pass
                
        tape_col_idx = None
        if len(all_values) >= 3:
            for i, val in enumerate(all_values[2]): 
                if str(val).strip() == "白貼":
                    tape_col_idx = i
                    break

        missing_in_db = []
        for r_idx, row in enumerate(all_values[3:], start=4): 
            if len(row) < 3: continue
            
            real_name = str(row[1]).strip() if len(row) > 1 else ""
            nickname = str(row[2]).strip() if len(row) > 2 else ""
            
            if nickname == "外賓人數" or real_name == "外賓人數":
                if sheet_exists and target_col_idx < len(row) and row[target_col_idx]:
                    try: guests_today = int(row[target_col_idx])
                    except: pass
                continue
            if nickname == "購買白貼" or real_name == "購買白貼":
                if sheet_exists and target_col_idx < len(row) and row[target_col_idx]:
                    try: purchased_tapes_today = int(row[target_col_idx])
                    except: pass
                continue

            if not nickname: continue # 綽號為必填，沒有就跳過

            # ✨ 組合顯示名稱：綽號(本名)
            display_name = f"{nickname}({real_name})" if real_name else nickname

            gender = str(row[3]).strip() if len(row) > 3 else ""
            if gender == "男": role = "🟦 底層"
            elif gender == "女": role = "🟥 上層"
            else: role = "未分類"

            if nickname not in db_map:
                missing_in_db.append([nickname, 0])
                db_map[nickname] = 0
            
            count = db_map[nickname]
            member_obj = {
                "nickname": nickname, 
                "real_name": real_name, 
                "display_name": display_name, 
                "role": role, 
                "row": r_idx, 
                "count": count
            }
            members.append(member_obj)
            
            if sheet_exists and target_col_idx < len(row) and str(row[target_col_idx]).strip() == "1":
                attendees_today.append(member_obj)
                
            if tape_col_idx and tape_col_idx < len(row) and str(row[tape_col_idx]).strip().upper() == "V":
                claimed_tapes.append(nickname)
                
        if missing_in_db:
            db_sheet.append_rows(missing_in_db)

        # 排序
        bases = [m["display_name"] for m in members if m["role"] == "🟦 底層"]
        flyers = [m["display_name"] for m in members if m["role"] == "🟥 上層"]
        bases.sort(key=lambda x: db_map.get(next(m['nickname'] for m in members if m['display_name']==x), 0), reverse=True)
        flyers.sort(key=lambda x: db_map.get(next(m['nickname'] for m in members if m['display_name']==x), 0), reverse=True)
        
        available_for_tape_objs = [m for m in members if m['nickname'] not in claimed_tapes]
        available_for_tape = [m['display_name'] for m in available_for_tape_objs]
        available_for_tape.sort(key=lambda x: db_map.get(next(m['nickname'] for m in members if m['display_name']==x), 0), reverse=True)

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
                    all_attendees_disp = selected_bases + selected_flyers
                    try:
                        if not sheet_exists:
                            valid_sheets = [s for s in sheet_names if re.match(r'^\d{6}$', s)]
                            valid_sheets.sort()
                            source_sheet = wb.worksheet(valid_sheets[-1]) if valid_sheets else sheet_for_read
                            
                            new_sheet = wb.duplicate_sheet(source_sheet.id, new_sheet_name=target_sheet_name)
                            days_in_month = calendar.monthrange(date_selected.year, date_selected.month)[1]
                            
                            # ✨ 換月清空邏輯：現在日期從 E欄(5) 到 AI欄(35)
                            cell_updates = []
                            for c_idx in range(5, 36): 
                                day = c_idx - 4
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

                        t_col = target_col_idx + 1 
                        updates = []
                        updates.append(gspread.Cell(row=1, col=t_col, value=st.session_state["current_user"]))
                        
                        name_to_row = {m['display_name']: m['row'] for m in members}
                        name_to_nickname = {m['display_name']: m['nickname'] for m in members}
                            
                        for disp_name in all_attendees_disp:
                            if disp_name in name_to_row:
                                updates.append(gspread.Cell(row=name_to_row[disp_name], col=t_col, value=1))
                                
                        if tape_col_idx:
                            t_tape_col = tape_col_idx + 1
                            for disp_name in selected_tapes:
                                if disp_name in name_to_row:
                                    updates.append(gspread.Cell(row=name_to_row[disp_name], col=t_tape_col, value="V"))
                        
                        # 處理外賓與購買
                        guest_r = next((i for i, row in enumerate(all_values) if len(row)>2 and (str(row[1]).strip()=="外賓人數" or str(row[2]).strip()=="外賓人數")), None)
                        if not guest_r:
                            guest_r = target_sheet.row_count + 1
                            target_sheet.update(values=[["x", "外賓人數", "外賓人數"]], range_name=f"A{guest_r}:C{guest_r}")
                        else: guest_r += 1 
                        updates.append(gspread.Cell(row=guest_r, col=t_col, value=guests_input_count if guests_input_count > 0 else ""))
                        
                        purchase_r = next((i for i, row in enumerate(all_values) if len(row)>2 and (str(row[1]).strip()=="購買白貼" or str(row[2]).strip()=="購買白貼")), None)
                        if not purchase_r:
                            purchase_r = target_sheet.row_count + 1
                            target_sheet.update(values=[["x", "購買白貼", "購買白貼"]], range_name=f"A{purchase_r}:C{purchase_r}")
                        else: purchase_r += 1
                        updates.append(gspread.Cell(row=purchase_r, col=t_col, value=tapes_purchased_input if tapes_purchased_input > 0 else ""))
                        
                        target_sheet.update_cells(updates)
                        
                        # 更新點名次數資料庫
                        db_all = db_sheet.get_all_records()
                        count_updates = []
                        for disp_name in all_attendees_disp:
                            nickname = name_to_nickname.get(disp_name)
                            for i, r in enumerate(db_all, start=2):
                                if str(r.get('姓名(綽號)', r.get('姓名', ''))) == nickname:
                                    count_updates.append(gspread.Cell(row=i, col=2, value=int(r.get('點名次數', 0)) + 1))
                                    break
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
            
            bases_present = [p['display_name'] for p in attendees_today if p['role'] == "🟦 底層"]
            flyers_present = [p['display_name'] for p in attendees_today if p['role'] == "🟥 上層"]
            
            st.markdown(f"**總出席隊員：{len(attendees_today)} 人**")
            st.markdown(f"🟦 **底層/男 ({len(bases_present)})：** " + "、".join(bases_present))
            st.markdown(f"🟥 **上層/女 ({len(flyers_present)})：** " + "、".join(flyers_present))
            
            col_s1, col_s2 = st.columns(2)
            col_s1.markdown(f"👤 **外賓人數：** `{guests_today}` 人")
            col_s2.markdown(f"📦 **購買白貼：** `{purchased_tapes_today}` 捲")
            
            st.write("---")
            claimed_tapes_disp = [m['display_name'] for m in members if m['nickname'] in claimed_tapes]
            st.markdown(f"🩹 **本月已領白貼人員 ({len(claimed_tapes_disp)})**")
            if claimed_tapes_disp:
                st.markdown("、".join(claimed_tapes_disp))
            else:
                st.caption("本月尚無人領取")
            
            st.write("---")
            with st.form("delete_form"):
                st.markdown("⚠️ **發現登記錯誤？在這裡取消**")
                to_delete_attend = st.multiselect("取消「點名」", options=[p['display_name'] for p in attendees_today])
                to_delete_tape = st.multiselect("取消本月「白貼領取」", options=claimed_tapes_disp)
                
                if st.form_submit_button("❌ 取消所選項並更新"):
                    if not to_delete_attend and not to_delete_tape:
                        st.warning("⚠️ 請先選擇要取消的項目！")
                    else:
                        try:
                            t_col = target_col_idx + 1
                            name_to_row = {m['display_name']: m['row'] for m in members}
                            del_updates = []
                            for disp_name in to_delete_attend:
                                del_updates.append(gspread.Cell(row=name_to_row[disp_name], col=t_col, value=""))
                            if tape_col_idx:
                                t_tape_col = tape_col_idx + 1
                                for disp_name in to_delete_tape:
                                    del_updates.append(gspread.Cell(row=name_to_row[disp_name], col=t_tape_col, value=""))
                                    
                            sheet_for_read.update_cells(del_updates)
                            st.session_state["success_msg"] = f"🗑️ 成功取消所選紀錄！"
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 取消時發生錯誤: {e}")

    # ================== 分頁 2：名單管理 ==================
    with tab_manage:
        st.header("⚙️ 隊員名單增刪管理")
        st.info("💡 姓名組合顯示為「綽號(本名)」。若只填寫綽號，則只顯示綽號。")
        col_add, col_edit, col_del = st.columns(3)
        
        with col_add:
            with st.form("add_member_form"):
                st.subheader("➕ 新增人員")
                new_nickname = st.text_input("輸入綽號 (必填，主要顯示用)")
                new_realname = st.text_input("輸入本名 (選填)")
                new_role = st.selectbox("選擇身分", ["🟦 底層", "🟥 上層"])
                
                if st.form_submit_button("新增至 G-Sheets"):
                    new_nickname = new_nickname.strip()
                    new_realname = new_realname.strip()
                    if not new_nickname:
                        st.warning("請至少輸入綽號！")
                    elif new_nickname in [m['nickname'] for m in members]:
                        st.warning("此綽號已存在！")
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
                            
                            # ✨ 寫入 B欄(本名), C欄(綽號), D欄(男/女)
                            gender_val = "男" if new_role == "🟦 底層" else "女"
                            sheet_for_read.update(values=[["x", new_realname, new_nickname, gender_val]], range_name=f"A{insert_idx}:D{insert_idx}")
                            
                            # 清空後面的打卡紀錄 (從 E 欄開始)
                            blanks = [[""] * 32]
                            sheet_for_read.update(values=blanks, range_name=f"E{insert_idx}:AJ{insert_idx}")
                            if tape_col_idx:
                                sheet_for_read.update_cell(insert_idx, tape_col_idx + 1, "")
                                
                            db_sheet.append_row([new_nickname, 0])
                            
                            disp = f"{new_nickname}({new_realname})" if new_realname else new_nickname
                            st.session_state["success_msg"] = f"✅ 成功將 {disp} 加入名單！"
                            st.rerun()
                        except Exception as e:
                            st.error(f"新增失敗：{e}")

        with col_edit:
            st.subheader("✏️ 修改資料")
            # ✨ 將選單拉出 Form 外，這樣選擇不同人時，下方的 Form 預設值就會自動更新！
            old_display_name = st.selectbox("選擇要修改的隊員", sorted([m['display_name'] for m in members]), key="edit_sel")
            selected_m = next((m for m in members if m['display_name'] == old_display_name), None)

            if selected_m:
                with st.form("edit_member_form"):
                    edit_new_nickname = st.text_input("輸入正確/新的綽號 (必填)", value=selected_m['nickname'])
                    edit_new_realname = st.text_input("輸入正確/新的本名 (選填)", value=selected_m['real_name'])
                    role_idx = 0 if selected_m['role'] == "🟦 底層" else (1 if selected_m['role'] == "🟥 上層" else 2)
                    edit_role = st.selectbox("修改身分", ["🟦 底層", "🟥 上層", "未分類"], index=role_idx)
                    
                    if st.form_submit_button("儲存修改"):
                        try:
                            target_nickname = edit_new_nickname.strip()
                            target_realname = edit_new_realname.strip()
                            
                            if not target_nickname:
                                st.warning("綽號不能為空喔！")
                            else:
                                gender_val = "男" if edit_role == "🟦 底層" else "女"
                                r = selected_m['row']
                                
                                # ✨ 一次性把新本名(B), 綽號(C), 性別(D) 寫入
                                sheet_for_read.update(values=[[target_realname, target_nickname, gender_val]], range_name=f"B{r}:D{r}")
                                    
                                # 若修改了綽號，同步更新資料庫
                                if target_nickname != selected_m['nickname']:
                                    cell = db_sheet.find(selected_m['nickname'], in_column=1)
                                    if cell:
                                        db_sheet.update_cell(cell.row, 1, target_nickname)
                                
                                st.session_state["success_msg"] = f"✏️ 成功更新「{target_nickname}」的資料！"
                                st.rerun()
                        except Exception as e:
                            st.error(f"修改失敗：{e}")

        with col_del:
            with st.form("delete_member_form"):
                st.subheader("🗑️ 刪除人員")
                st.error("刪除後將完全移除！")
                del_display_name = st.selectbox("選擇要永久刪除的隊員", sorted([m['display_name'] for m in members]))
                
                if st.form_submit_button("永久刪除該隊員"):
                    try:
                        del_m = next(m for m in members if m['display_name'] == del_display_name)
                        sheet_for_read.delete_rows(del_m['row'])
                        
                        cell = db_sheet.find(del_m['nickname'], in_column=1)
                        if cell: db_sheet.delete_rows(cell.row)
                        
                        st.session_state["success_msg"] = f"🗑️ 已將 {del_display_name} 徹底移除！"
                        st.rerun()
                    except Exception as e:
                        st.error(f"刪除失敗：{e}")

if __name__ == "__main__":
    main()
