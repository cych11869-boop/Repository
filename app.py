import streamlit as st
import pandas as pd
import io
import random
from ortools.sat.python import cp_model

st.set_page_config(page_title="醫療 AI 排班系統", page_icon="🏥", layout="wide")

st.title("🏥 醫療 AI 智能排班系統 V4.1 (固定座位鎖定版)")
st.markdown("上傳您的排班底稿，AI 將在 60 秒內為您產出完美班表！")

def generate_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({
            '參數名稱': ['排班週數 (Weeks)', 'N班最低包班天數', 'E班最低包班天數'],
            '設定值': [4, 12, 12],
            '說明': ['1~4週', '整個週期內N班最少要上幾天', '整個週期內E班最少要上幾天']
        }).to_excel(writer, sheet_name='基本設定', index=False)
        
        grid_data = [
            ['[需求]N班', 'REQ'] + ['']*7 + [2]*28,
            ['[需求]E班', 'REQ'] + ['']*7 + [3]*28,
            ['[需求]D班', 'REQ'] + ['']*7 + [5]*28, 
        ]
        staff_list = [('R01', 'N'), ('R02', 'N'), ('R03', 'E'), ('R04', 'R'), ('R15_孕婦', 'D')]
        columns = ['員工代碼', '群組(N/E/R/D)'] + [f'歷史_{i}' for i in range(1, 8)] + [f'Day_{i}' for i in range(1, 29)]
        for e, grp in staff_list: grid_data.append([e, grp] + ['X']*7 + ['']*28)
        pd.DataFrame(grid_data, columns=columns).to_excel(writer, sheet_name='排班底稿', index=False)

        dict_data = {
            '系統代碼': ['N1', 'N2', 'E1', 'E2', 'E3', 'D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7'],
            '平日顯示名稱': ['N1 0000-0800', 'N2 0000-0800', 'E1 1600-0000', 'E2 1600-0000', 'E3 1600-0000', 'D1 0730-1600', 'DBR 0730-1600', 'DBA 0730-1600', 'DCB 0800-1700', 'DF 0800-1700', 'DF 0800-1700', 'DF 0800-1700'],
            '星期六顯示名稱': ['N1 0000-0800', 'N2 0000-0800', 'E1 1600-0000', 'E2 1600-0000', 'E3 1600-0000', 'D6 0730-1600', 'DBR 0800-1600', 'DBA 0800-1600', 'D4 0730-1600', 'D8 0730-1600', 'D9 0800-1600', 'DF 0800-1700'],
            '星期日顯示名稱': ['N1 0000-0800', 'N2 0000-0800', 'E1 1600-0000', 'E2 1600-0000', 'E3 1600-0000', 'D4 0800-1600', 'DBR 0800-1600', 'DBA 0800-1600', 'DF 0800-1700', 'DF 0800-1700', 'DF 0800-1700', 'DF 0800-1700']
        }
        pd.DataFrame(dict_data).to_excel(writer, sheet_name='班別設定', index=False)
    return output.getvalue()

with st.expander("第一次使用？點此下載空白模板"):
    st.markdown("模板已升級！內含**【班別設定】**分頁，可自由修改貴單位專屬的上下班時間與代號。")
    st.download_button("📥 下載 V4 升級版模板", data=generate_template(), file_name="Schedule_Template_V4.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.sidebar.header("⚙️ 參數設定")
st.sidebar.markdown("網頁版可直接在此拉動參數，**將覆蓋 Excel 內的設定**。")
weeks = st.sidebar.slider("排班週數 (Weeks)", 1, 4, 1)
n_min_shifts = st.sidebar.slider("N班最低包班天數", 0, 20, 0)
e_min_shifts = st.sidebar.slider("E班最低包班天數", 0, 20, 3)

uploaded_file = st.file_uploader("📂 請上傳排班底稿", type=['xlsx'])

if uploaded_file is not None:
    if st.button("🚀 啟動 AI 最佳化排班", use_container_width=True):
        with st.spinner("🧠 AI 正在讀取並進行深度運算... (約需 60 秒)"):
            try:
                dict_weekday, dict_sat, dict_sun = {}, {}, {}
                try:
                    df_dict = pd.read_excel(uploaded_file, sheet_name='班別設定')
                    for _, row in df_dict.iterrows():
                        code = str(row['系統代碼']).strip()
                        dict_weekday[code] = str(row['平日顯示名稱']).strip()
                        dict_sat[code] = str(row['星期六顯示名稱']).strip()
                        dict_sun[code] = str(row['星期日顯示名稱']).strip()
                except ValueError:
                    dict_weekday = {'N1': 'N1 0000-0800', 'N2': 'N2 0000-0800', 'E1': 'E1 1600-0000', 'E2': 'E2 1600-0000', 'E3': 'E3 1600-0000', 'D1': 'D1 0730-1600', 'D2': 'DBR 0730-1600', 'D3': 'DBA 0730-1600', 'D4': 'DCB 0800-1700', 'D5': 'DF 0800-1700', 'D6': 'DF 0800-1700', 'D7': 'DF 0800-1700'}
                    dict_sat = {'N1': 'N1 0000-0800', 'N2': 'N2 0000-0800', 'E1': 'E1 1600-0000', 'E2': 'E2 1600-0000', 'E3': 'E3 1600-0000', 'D1': 'D6 0730-1600', 'D2': 'DBR 0800-1600', 'D3': 'DBA 0800-1600', 'D4': 'D4 0730-1600', 'D5': 'D8 0730-1600', 'D6': 'D9 0800-1600'}
                    dict_sun = {'N1': 'N1 0000-0800', 'N2': 'N2 0000-0800', 'E1': 'E1 1600-0000', 'E2': 'E2 1600-0000', 'E3': 'E3 1600-0000', 'D1': 'D4 0800-1600', 'D2': 'DBR 0800-1600', 'D3': 'DBA 0800-1600'}

                target_days, num_days, total_x_quota = weeks * 7, 7 + weeks * 7, weeks * 2
                df_grid = pd.read_excel(uploaded_file, sheet_name='排班底稿')
                
                staff_n, staff_e, staff_r, staff_d = [], [], [], []
                history_map, locked_shifts = {}, []
                
                # 🌟 V4.1 新增：紀錄精確的座位鎖定 (如 D1, D2, N1)
                specific_locks = {} 
                
                req_n = {d: 2 for d in range(7, num_days)}
                req_e = {d: 3 for d in range(7, num_days)}
                req_d = {d: 5 for d in range(7, num_days)}

                for _, row in df_grid.iterrows():
                    e = str(row['員工代碼']).strip()
                    grp = str(row.iloc[1]).strip().upper()
                    
                    if grp == 'REQ':
                        for i in range(1, target_days + 1):
                            val = row[f'Day_{i}']
                            if pd.notna(val) and str(val).strip() != '':
                                v = int(float(val))
                                if 'N班' in e: req_n[6+i] = v
                                elif 'E班' in e: req_e[6+i] = v
                                elif 'D班' in e: req_d[6+i] = v
                        continue

                    if grp == 'N': staff_n.append(e)
                    elif grp == 'E': staff_e.append(e)
                    elif grp == 'D': staff_d.append(e)
                    else: staff_r.append(e)

                    history_map[e] = [str(row[f'歷史_{i}']).strip().upper() for i in range(1, 8)]
                    
                    for i in range(1, target_days + 1):
                        val = row[f'Day_{i}']
                        if pd.notna(val) and str(val).strip() != '':
                            raw_val = str(val).strip().upper()
                            
                            # 🌟 V4.1 核心解析：擷取大分類給 AI，保留細項給 Python
                            major_val = raw_val
                            if raw_val.startswith('D') and raw_val != 'D': major_val = 'D'
                            elif raw_val.startswith('N') and raw_val != 'N': major_val = 'N'
                            elif raw_val.startswith('E') and raw_val != 'E': major_val = 'E'
                            
                            locked_shifts.append((e, 6 + i, major_val))
                            
                            # 如果使用者有指定細項 (例如輸入了 D2 而不是 D)，則存入特定字典
                            if major_val != raw_val:
                                specific_locks[(e, 6 + i)] = raw_val

                all_staff = staff_n + staff_e + staff_r + staff_d
                shift_codes = {'X': 1, 'F': 2, 'N': 3, 'E': 4, 'D': 5}
                all_shifts = list(shift_codes.values())

                model = cp_model.CpModel()
                works = {}
                for e in all_staff:
                    for d in range(num_days):
                        for s in all_shifts: works[(e, d, s)] = model.NewBoolVar(f'w_{e}_{d}_{s}')
                        model.AddExactlyOne(works[(e, d, s)] for s in all_shifts)

                for e in all_staff:
                    # 歷史紀錄處理：如果歷史紀錄是 D1，只抓第一個字 D 放進系統
                    for d in range(7): 
                        hist_major = history_map[e][d][0] if history_map[e][d].startswith(('D','N','E')) else history_map[e][d]
                        model.Add(works[(e, d, shift_codes.get(hist_major, 1))] == 1)

                for d in range(7, num_days):
                    model.Add(sum(works[(e, d, shift_codes['N'])] for e in all_staff) == req_n[d])
                    model.Add(sum(works[(e, d, shift_codes['E'])] for e in all_staff) == req_e[d])
                    is_weekend = (d % 7 == 5) or (d % 7 == 6)
                    if is_weekend: model.Add(sum(works[(e, d, shift_codes['D'])] for e in all_staff) == req_d[d])
                    else: model.Add(sum(works[(e, d, shift_codes['D'])] for e in all_staff) >= req_d[d])

                for e in all_staff:
                    for d in range(num_days - 6): model.Add(sum(works[(e, d+i, 1)] for i in range(7)) >= 1)
                    for d in range(1, num_days): model.Add(works[(e, d-1, 3)] + works[(e, d-1, 1)] == 1).OnlyEnforceIf(works[(e, d, 3)])
                    for d in range(num_days - 1): model.Add(works[(e, d+1, 4)] + works[(e, d+1, 1)] == 1).OnlyEnforceIf(works[(e, d, 4)])

                for e in all_staff: model.Add(sum(works[(e, d, 1)] for d in range(7, num_days)) == total_x_quota) 
                
                for e in staff_e:
                    for d in range(7, num_days): model.Add(works[(e, d, shift_codes['N'])] == 0)
                for e in staff_n:
                    for d in range(7, num_days): model.Add(works[(e, d, shift_codes['E'])] == 0)
                for e in staff_d:
                    for d in range(7, num_days):
                        model.Add(works[(e, d, shift_codes['N'])] == 0)
                        model.Add(works[(e, d, shift_codes['E'])] == 0)

                for e in staff_n: model.Add(sum(works[(e, d, shift_codes['N'])] for d in range(7, num_days)) >= n_min_shifts)
                for e in staff_e: model.Add(sum(works[(e, d, shift_codes['E'])] for d in range(7, num_days)) >= e_min_shifts)

                locked_f_coords = {(e, d) for e, d, s in locked_shifts if s == 'F'}
                for e, d, s in locked_shifts: model.Add(works[(e, d, shift_codes[s])] == 1)
                for e in all_staff:
                    for d in range(7, num_days):
                        if (e, d) not in locked_f_coords: model.Add(works[(e, d, shift_codes['F'])] == 0)

                penalties = []
                is_off = {}
                for e in all_staff:
                    for d in range(num_days): is_off[(e, d)] = works[(e, d, shift_codes['X'])] + works[(e, d, shift_codes['F'])]
                
                weekend_days = [d for d in range(7, num_days) if d % 7 in (5, 6)]
                min_wknd_off = max(1, weeks // 2)      
                max_wknd_off = min_wknd_off + 2        
                
                for e in all_staff:
                    weekend_offs = sum(works[(e, d, shift_codes['X'])] + works[(e, d, shift_codes['F'])] for d in weekend_days)
                    excess_wknd = model.NewIntVar(0, len(weekend_days), f'exc_wknd_{e}')
                    model.Add(excess_wknd >= weekend_offs - max_wknd_off)
                    penalties.append(excess_wknd * 5) 
                    short_wknd = model.NewIntVar(0, len(weekend_days), f'sho_wknd_{e}')
                    model.Add(short_wknd >= min_wknd_off - weekend_offs)
                    penalties.append(short_wknd * 5)  

                    for d in range(1, num_days - 1):
                        iso_work = model.NewBoolVar(f'iw_{e}_{d}')
                        model.Add(iso_work >= is_off[(e, d-1)] + (1 - is_off[(e, d)]) + is_off[(e, d+1)] - 2)
                        penalties.append(iso_work * 2)
                        iso_off = model.NewBoolVar(f'io_{e}_{d}')
                        model.Add(iso_off >= (1 - is_off[(e, d-1)]) + is_off[(e, d)] + (1 - is_off[(e, d+1)]) - 2)
                        penalties.append(iso_off * 1)

                    for d in range(7, num_days - 1):
                        if d % 7 == 4:
                            free_ride_fri = model.NewBoolVar(f'fr_fri_{e}_{d}')
                            model.Add(free_ride_fri >= works[(e, d, shift_codes['F'])] + works[(e, d+1, shift_codes['X'])] - 1)
                            penalties.append(free_ride_fri * 15)
                    for d in range(8, num_days):
                        if d % 7 == 0:
                            free_ride_mon = model.NewBoolVar(f'fr_mon_{e}_{d}')
                            model.Add(free_ride_mon >= works[(e, d, shift_codes['F'])] + works[(e, d-1, shift_codes['X'])] - 1)
                            penalties.append(free_ride_mon * 15)

                weekday_days = [d for d in range(7, num_days) if d % 7 not in (5, 6)]
                mid_week_days = [d for d in weekday_days if d % 7 in (1, 2, 3)] 
                mon_fri_days = [d for d in weekday_days if d % 7 in (0, 4)]     

                for d in mid_week_days:
                    d_count = sum(works[(e, d, shift_codes['D'])] for e in all_staff)
                    extra_d = model.NewIntVar(0, len(all_staff), f'extra_d_{d}')
                    model.Add(extra_d == d_count - req_d[d])
                    penalties.append(extra_d * 20) 

                if mon_fri_days:
                    max_mf = model.NewIntVar(0, len(all_staff), 'max_mf')
                    min_mf = model.NewIntVar(0, len(all_staff), 'min_mf')
                    for d in mon_fri_days:
                        d_count = sum(works[(e, d, shift_codes['D'])] for e in all_staff)
                        extra_d = model.NewIntVar(0, len(all_staff), f'extra_d_{d}')
                        model.Add(extra_d == d_count - req_d[d])
                        model.Add(max_mf >= extra_d)
                        model.Add(min_mf <= extra_d)
                    penalties.append((max_mf - min_mf) * 10)

                model.Minimize(sum(penalties))

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 60.0 
                status = solver.Solve(model)

                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✅ 運算完成！已為您自動分配固定座位。")
                    reverse_shift_codes = {1: 'X', 2: 'F', 3: 'N', 4: 'E', 5: 'D'}
                    schedule_data = []
                    for e in all_staff:
                        schedule_data.append([e] + [reverse_shift_codes[s] for d in range(7, num_days) for s in all_shifts if solver.Value(works[(e, d, s)]) == 1])
                    df_out = pd.DataFrame(schedule_data, columns=['Staff'] + [f'Day_{i}' for i in range(1, target_days + 1)])
                    
                    # ==========================================================
                    # 🌟 V4.1 座位分派邏輯：優先滿足特定鎖定，剩餘才洗牌分配
                    # ==========================================================
                    for col in df_out.columns[1:]:
                        day_idx = int(col.split('_')[1])
                        d = 6 + day_idx
                        is_sat = (d % 7 == 5)
                        is_sun = (d % 7 == 6)
                        
                        # 處理 N, E, D 的派位
                        for shift_type in ['N', 'E', 'D']:
                            idx_list = df_out[df_out[col] == shift_type].index.tolist()
                            if not idx_list: continue
                            
                            # 生成該班別的座位池 (例如 ['D1', 'D2', 'D3', 'D4', 'D5'])
                            seat_pool = [f'{shift_type}{i}' for i in range(1, len(idx_list) + 1)]
                            
                            remaining_idx = []
                            
                            # 【步驟一】優先派發「手動指定」的座位
                            for idx in idx_list:
                                emp_id = df_out.at[idx, 'Staff']
                                specific_seat = specific_locks.get((emp_id, d))
                                
                                # 如果這個人有指定座位，且屬於目前的班別 (例如 D)
                                if specific_seat and specific_seat.startswith(shift_type):
                                    # 直接給他指定的座位，並從座位池裡拿掉一個位置保持平衡
                                    df_out.at[idx, col] = specific_seat
                                    if specific_seat in seat_pool:
                                        seat_pool.remove(specific_seat)
                                    else:
                                        seat_pool.pop() # 如果他指定的座位太特別(如D8)，隨便抽掉一個維持名額
                                else:
                                    remaining_idx.append(idx)
                                    
                            # 【步驟二】把剩下的座位洗牌，發給剩下沒有指定座位的人
                            random.shuffle(seat_pool)
                            for i, idx in enumerate(remaining_idx):
                                df_out.at[idx, col] = seat_pool[i]

                        # 進行字典翻譯
                        for idx in df_out.index:
                            code = df_out.at[idx, col]
                            if code in ['X', 'F'] or 'X' in code: continue
                            
                            fallback = f"{code} 0800-1700" if code.startswith('D') else f"{code} 時間未定"
                            
                            if code.startswith('N') or code.startswith('E'):
                                df_out.at[idx, col] = dict_weekday.get(code, code)
                            else: # D 班
                                if is_sat: df_out.at[idx, col] = dict_sat.get(code, fallback)
                                elif is_sun: df_out.at[idx, col] = dict_sun.get(code, fallback)
                                else: df_out.at[idx, col] = dict_weekday.get(code, fallback)

                    # 處理一例一休
                    x_format_map = {1: '休', 2: '例'}
                    for idx in df_out.index:
                        for w in range(weeks):
                            c_X = 1
                            for col in [f'Day_{d}' for d in range(w*7 + 1, w*7 + 8)]:
                                if df_out.at[idx, col] == 'X':
                                    df_out.at[idx, col] = x_format_map.get(c_X, f'X{c_X}')
                                    c_X += 1
                    
                    st.dataframe(df_out)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_out.to_excel(writer, index=False, sheet_name='神仙班表')
                    
                    st.download_button(
                        label="📥 下載神仙班表.xlsx",
                        data=output.getvalue(),
                        file_name="Schedule_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.balloons()
                else:
                    st.error("❌ 依然排不出來！請檢查 Excel 的請假是否產生矛盾，或人力需求是否過高。")
            except Exception as e:
                st.error(f"執行時發生錯誤：{e}")
