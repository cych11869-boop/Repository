# Repository
排班
!pip install ortools pandas openpyxl

import os
import random
import pandas as pd
from ortools.sat.python import cp_model

INPUT_FILE = 'Schedule_Input.xlsx'
OUTPUT_FILE = 'Schedule_Output.xlsx'

# =====================================================================
# 📝 第一階段：建立標準化輸入模板
# =====================================================================
def create_excel_template():
    df_settings = pd.DataFrame({
        '參數名稱': ['排班週數 (Weeks)', 'N班最低包班天數', 'E班最低包班天數'],
        '設定值': [4, 12, 12],
        '說明': ['1~4週 (若採滾動排班，請維持設4週，讓AI向未來借假，最後只取前1週)', '整個週期內N班最少要上幾天', '整個週期內E班最少要上幾天']
    })

    # [動態人力矩陣]：允許每日針對不同班別設定具體需求人數
    grid_data = [
        ['[需求]N班', 'REQ'] + ['']*7 + [2]*28,
        ['[需求]E班', 'REQ'] + ['']*7 + [3]*28,
        ['[需求]D班', 'REQ'] + ['']*7 + [5]*28,
    ]

    # [群組定義] N(包大夜)、E(包小夜)、R(一般輪班)、D(純白班/孕婦/育嬰)
    staff_list = [
        ('R01', 'N'), ('R02', 'N'), ('R12', 'N'),
        ('R03', 'E'), ('R05', 'E'), ('R06', 'E'), ('R07', 'E'), ('R13', 'E'),
        ('R04', 'R'), ('R08', 'R'), ('R09', 'R'), ('R10', 'R'), ('R11', 'R'), ('R14', 'R'),
        ('R15_孕婦', 'D')
    ]
    columns = ['員工代碼', '群組(N/E/R/D)'] + [f'歷史_{i}' for i in range(1, 8)] + [f'Day_{i}' for i in range(1, 29)]

    for e, grp in staff_list:
        grid_data.append([e, grp] + ['X']*7 + ['']*28)

    with pd.ExcelWriter(INPUT_FILE, engine='openpyxl') as writer:
        df_settings.to_excel(writer, sheet_name='基本設定', index=False)
        pd.DataFrame(grid_data, columns=columns).to_excel(writer, sheet_name='排班底稿', index=False)
    print(f"🎉 模板已生成：【{INPUT_FILE}】")

# =====================================================================
# 🚨 第二階段：事前診斷雷達 (預防 AI 遇到死結直接崩潰)
# =====================================================================
def pre_flight_check(history_map, locked_shifts, all_staff, target_days, req_n, req_e, req_d, staff_n, staff_e, staff_d):
    print("\n🕵️‍♂️ [事前診斷雷達] 正在核對動態人力與接班邏輯...")
    errors = 0
    known = {}

    # 建立已知日曆，包含歷史紀錄與手動鎖定的班 (F/D/E/N/X)
    for e in all_staff:
        for d in range(7): known[(e, d)] = history_map[e][d]
    for e, d, s in locked_shifts: known[(e, d)] = s

    for e in all_staff:
        for d in range(1, 7 + target_days):
            s_yest = known.get((e, d-1))
            s_today = known.get((e, d))

            # [防呆] N班防護罩：大夜班的前一天只能是大夜(N)或排休(X)
            if s_today == 'N' and s_yest not in ['N', 'X', None]:
                print(f"❌ 【接班衝突】{e} (Day_{d-6}) 鎖定 N，但前一天是 {s_yest}。"); errors += 1
            # [防呆] E班防護罩：小夜班的隔天不能排白班或大夜，因為休息時間不足
            if s_yest == 'E' and s_today not in ['E', 'X', None]:
                print(f"❌ 【接班衝突】{e} (Day_{d-7}) 是 E，但隔天是 {s_today}。"); errors += 1

        # [防呆] 連續上班 7 天檢查 (注意：F 被視為出勤，不具備斷班功能)
        for start_d in range(7 + target_days - 6):
            if all(known.get((e, start_d + i)) not in ['X', None] for i in range(7)):
                print(f"❌ 【連七崩潰】{e} 發生連續7天無 X 假(F不等於休)。"); errors += 1

    # [防呆] 群組與班別衝突檢查
    for e, d, s in locked_shifts:
        if e in staff_n and s == 'E': print(f"❌ 【群組衝突】{e} 是 N 包班，不可鎖定 E。"); errors += 1
        if e in staff_e and s == 'N': print(f"❌ 【群組衝突】{e} 是 E 包班，不可鎖定 N。"); errors += 1
        if e in staff_d and s in ['N', 'E']: print(f"❌ 【群組衝突】{e} 是純白班，不可鎖定 {s} 班。"); errors += 1

    # [防呆] 單日可用戰力檢查：當天畫 F 或 X 的人數過多，導致無法滿足該日最低需求
    for d in range(7, 7 + target_days):
        locked_out = sum(1 for e in all_staff if known.get((e, d)) in ['F', 'X'])
        avail = len(all_staff) - locked_out
        req_total = req_n[d] + req_e[d] + req_d[d]
        if avail < req_total:
            print(f"❌ 【人力斷層】Day_{d-6} 剩 {avail} 人，但需求高達 N({req_n[d]})+E({req_e[d]})+D({req_d[d]})={req_total} 人。")
            errors += 1

    if errors == 0:
        print("✅ 診斷通過！人力設定皆合法。\n")
        return True
    return False

# =====================================================================
# 🧠 第三階段：AI 核心排班演算引擎
# =====================================================================
def run_schedule_optimizer():
    print(f"📥 正在讀取 {INPUT_FILE} ...")
    df_settings = pd.read_excel(INPUT_FILE, sheet_name='基本設定')
    weeks = int(df_settings.loc[0, '設定值'])
    n_min_shifts, e_min_shifts = int(df_settings.loc[1, '設定值']), int(df_settings.loc[2, '設定值'])

    # 計算時間窗：目標天數(如28天)、總天數(含歷史7天)、法規排休總配額(每週2天X)
    target_days, num_days, total_x_quota = weeks * 7, 7 + weeks * 7, weeks * 2

    df_grid = pd.read_excel(INPUT_FILE, sheet_name='排班底稿')
    staff_n, staff_e, staff_r, staff_d = [], [], [], []
    history_map, locked_shifts = {}, []

    # 初始化每日動態人力預設值
    req_n = {d: 2 for d in range(7, num_days)}
    req_e = {d: 3 for d in range(7, num_days)}
    req_d = {d: 5 for d in range(7, num_days)}

    # 讀取 Excel 資料
    for _, row in df_grid.iterrows():
        e = str(row['員工代碼']).strip()
        grp = str(row.iloc[1]).strip().upper()

        # 抓取並覆蓋每日動態需求
        if grp == 'REQ':
            for i in range(1, target_days + 1):
                val = row[f'Day_{i}']
                if pd.notna(val) and str(val).strip() != '':
                    v = int(float(val))
                    if 'N班' in e: req_n[6+i] = v
                    elif 'E班' in e: req_e[6+i] = v
                    elif 'D班' in e: req_d[6+i] = v
            continue

        # 員工群組分類
        if grp == 'N': staff_n.append(e)
        elif grp == 'E': staff_e.append(e)
        elif grp == 'D': staff_d.append(e)
        else: staff_r.append(e)

        history_map[e] = [str(row[f'歷史_{i}']).strip().upper() for i in range(1, 8)]
        for i in range(1, target_days + 1):
            val = row[f'Day_{i}']
            if pd.notna(val) and str(val).strip() != '':
                locked_shifts.append((e, 6 + i, str(val).strip().upper()))

    all_staff = staff_n + staff_e + staff_r + staff_d
    shift_codes = {'X': 1, 'F': 2, 'N': 3, 'E': 4, 'D': 5}
    all_shifts = list(shift_codes.values())

    if not pre_flight_check(history_map, locked_shifts, all_staff, target_days, req_n, req_e, req_d, staff_n, staff_e, staff_d):
        return

    # [核心] 建立 Google OR-Tools 數學模型
    model = cp_model.CpModel()
    works = {}
    for e in all_staff:
        for d in range(num_days):
            for s in all_shifts: works[(e, d, s)] = model.NewBoolVar(f'w_{e}_{d}_{s}')
            # 絕對法則：每人每天只能剛好上一種班(或休假)
            model.AddExactlyOne(works[(e, d, s)] for s in all_shifts)

    # 鎖定歷史紀錄 7 天
    for e in all_staff:
        for d in range(7): model.Add(works[(e, d, shift_codes.get(history_map[e][d], 1))] == 1)

    # ---------------------------------------------------------
    # 【邏輯 1】每日動態人力與假日鎖定機制
    # ---------------------------------------------------------
    for d in range(7, num_days):
        model.Add(sum(works[(e, d, shift_codes['N'])] for e in all_staff) == req_n[d])
        model.Add(sum(works[(e, d, shift_codes['E'])] for e in all_staff) == req_e[d])
        is_weekend = (d % 7 == 5) or (d % 7 == 6)
        if is_weekend:
            # 假日：D班人數「絕對等於」Excel需求，避免 AI 把多餘上班配額塞在假日
            model.Add(sum(works[(e, d, shift_codes['D'])] for e in all_staff) == req_d[d])
        else:
            # 平日：D班人數「最少等於」Excel需求，多出來的總上班配額會流動到平日吸收
            model.Add(sum(works[(e, d, shift_codes['D'])] for e in all_staff) >= req_d[d])

    # ---------------------------------------------------------
    # 【邏輯 2】臨床防呆與接班法規
    # ---------------------------------------------------------
    for e in all_staff:
        # 防連七：任意連續 7 天內，[代碼1 (也就是X)] 的加總至少要 1 天。F不具備斷班效果。
        for d in range(num_days - 6): model.Add(sum(works[(e, d+i, 1)] for i in range(7)) >= 1)
        # N 班防護：如果今天是 N，昨天只能是 N 或是 X。
        for d in range(1, num_days): model.Add(works[(e, d-1, 3)] + works[(e, d-1, 1)] == 1).OnlyEnforceIf(works[(e, d, 3)])
        # E 班防護：如果明天是 E，今天只能是 E 或是 X。(反向思考，確保 E 班隔天不能接 D 或 N)
        for d in range(num_days - 1): model.Add(works[(e, d+1, 4)] + works[(e, d+1, 1)] == 1).OnlyEnforceIf(works[(e, d, 4)])

    # [配額脫鉤] 只計算 X 作為排休額度，完全不計算 F
    for e in all_staff: model.Add(sum(works[(e, d, 1)] for d in range(7, num_days)) == total_x_quota)

    # 群組底線限制
    for e in staff_e:
        for d in range(7, num_days): model.Add(works[(e, d, shift_codes['N'])] == 0)
    for e in staff_n:
        for d in range(7, num_days): model.Add(works[(e, d, shift_codes['E'])] == 0)
    for e in staff_d: # 純白班/孕婦群組防護：絕對不上 N 與 E
        for d in range(7, num_days):
            model.Add(works[(e, d, shift_codes['N'])] == 0)
            model.Add(works[(e, d, shift_codes['E'])] == 0)

    for e in staff_n: model.Add(sum(works[(e, d, shift_codes['N'])] for d in range(7, num_days)) >= n_min_shifts)
    for e in staff_e: model.Add(sum(works[(e, d, shift_codes['E'])] for d in range(7, num_days)) >= e_min_shifts)

    # 鎖定預先畫好的假與班，並確保沒畫 F 的格子不會無故生出 F
    locked_f_coords = {(e, d) for e, d, s in locked_shifts if s == 'F'}
    for e, d, s in locked_shifts: model.Add(works[(e, d, shift_codes[s])] == 1)
    for e in all_staff:
        for d in range(7, num_days):
            if (e, d) not in locked_f_coords: model.Add(works[(e, d, shift_codes['F'])] == 0)


    # =================================================================
    # ⚖️ 第四階段：軟性約束 (扣分懲罰機制)，這決定了班表的「完美度」
    # =================================================================
    penalties = []

    # 先建立一個布林字典 is_off：只要這天是 X 或 F，都算是「實質上沒在醫院」
    is_off = {}
    for e in all_staff:
        for d in range(num_days): is_off[(e, d)] = works[(e, d, shift_codes['X'])] + works[(e, d, shift_codes['F'])]

    # ---------------------------------------------------------
    # 【扣分邏輯 A】假日公平分配機制 (防老鳥狂休六日)
    # ---------------------------------------------------------
    weekend_days = [d for d in range(7, num_days) if d % 7 in (5, 6)]
    min_wknd_off = max(1, weeks // 2)      # 算出最少該休幾個假日 (例如4週配2天)
    max_wknd_off = min_wknd_off + 2        # 算出最多能休幾個假日 (例如4週配4天)

    for e in all_staff:
        weekend_offs = sum(works[(e, d, shift_codes['X'])] + works[(e, d, shift_codes['F'])] for d in weekend_days)
        # 休太多假日，每多一天重扣 5 分
        excess_wknd = model.NewIntVar(0, len(weekend_days), f'exc_wknd_{e}')
        model.Add(excess_wknd >= weekend_offs - max_wknd_off)
        penalties.append(excess_wknd * 5)
        # 休太少假日，每少一天重扣 5 分
        short_wknd = model.NewIntVar(0, len(weekend_days), f'sho_wknd_{e}')
        model.Add(short_wknd >= min_wknd_off - weekend_offs)
        penalties.append(short_wknd * 5)

        # ---------------------------------------------------------
        # 【扣分邏輯 B】碎班最佳化 (防做一休一)
        # ---------------------------------------------------------
        for d in range(1, num_days - 1):
            # 判斷孤立上班日：休(d-1) - 班(d) - 休(d+1)
            iso_work = model.NewBoolVar(f'iw_{e}_{d}')
            model.Add(iso_work >= is_off[(e, d-1)] + (1 - is_off[(e, d)]) + is_off[(e, d+1)] - 2)
            penalties.append(iso_work * 2) # 最討厭做一休一，扣 2 分

            # 判斷孤立休假日：班(d-1) - 休(d) - 班(d+1)
            iso_off = model.NewBoolVar(f'io_{e}_{d}')
            model.Add(iso_off >= (1 - is_off[(e, d-1)]) + is_off[(e, d)] + (1 - is_off[(e, d+1)]) - 2)
            penalties.append(iso_off * 1)  # 休一做一，扣 1 分

    # ---------------------------------------------------------
    # 【扣分邏輯 C】搭便車防堵機制 (Anti-Free-Rider)
    # ---------------------------------------------------------
    for e in all_staff:
        # 防堵利用「週五請特休F」吸附「週六排休X」來湊連假
        for d in range(7, num_days - 1):
            if d % 7 == 4: # 4代表週五
                free_ride_fri = model.NewBoolVar(f'fr_fri_{e}_{d}')
                model.Add(free_ride_fri >= works[(e, d, shift_codes['F'])] + works[(e, d+1, shift_codes['X'])] - 1)
                penalties.append(free_ride_fri * 15) # 抓到重扣 15 分，AI寧可讓他週六來上班

        # 防堵利用「週一請特休F」吸附「週日排休X」
        for d in range(8, num_days):
            if d % 7 == 0: # 0代表週一
                free_ride_mon = model.NewBoolVar(f'fr_mon_{e}_{d}')
                model.Add(free_ride_mon >= works[(e, d, shift_codes['F'])] + works[(e, d-1, shift_codes['X'])] - 1)
                penalties.append(free_ride_mon * 15)

    # ---------------------------------------------------------
    # 【扣分邏輯 D】平日多餘人力引流與水庫攤平 (Minimax Routing)
    # ---------------------------------------------------------
    weekday_days = [d for d in range(7, num_days) if d % 7 not in (5, 6)]
    mid_week_days = [d for d in weekday_days if d % 7 in (1, 2, 3)] # 週二、週三、週四冷門區
    mon_fri_days = [d for d in weekday_days if d % 7 in (0, 4)]     # 週一、週五熱門區

    # 1. 嚴格壓抑二三四：只要超出最低需求，每多一人重扣 20 分！(把多餘人力擠出冷門區)
    for d in mid_week_days:
        d_count = sum(works[(e, d, shift_codes['D'])] for e in all_staff)
        extra_d = model.NewIntVar(0, len(all_staff), f'extra_d_{d}')
        model.Add(extra_d == d_count - req_d[d])
        penalties.append(extra_d * 20)

    # 2. 鼓勵週一週五吸收多餘人力，但必須互相攤平 (Minimax)
    if mon_fri_days:
        max_mf = model.NewIntVar(0, len(all_staff), 'max_mf')
        min_mf = model.NewIntVar(0, len(all_staff), 'min_mf')
        for d in mon_fri_days:
            d_count = sum(works[(e, d, shift_codes['D'])] for e in all_staff)
            extra_d = model.NewIntVar(0, len(all_staff), f'extra_d_{d}')
            model.Add(extra_d == d_count - req_d[d])
            model.Add(max_mf >= extra_d) # 找出週一或週五，額外人力最多的一天
            model.Add(min_mf <= extra_d) # 找出最少的一天

        # 將「最大落差」當作扣分，逼迫 AI 將多餘人力「公平、平均」地分給週一跟週五
        penalties.append((max_mf - min_mf) * 10)

    # 目標：將所有懲罰分數降到最低
    model.Minimize(sum(penalties))

    print("🧠 AI 正在進行深度引流與最佳化... (約需 60 秒)")
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        print(f"✅ 計算完成！")
        reverse_shift_codes = {1: 'X', 2: 'F', 3: 'N', 4: 'E', 5: 'D'}
        schedule_data = []
        for e in all_staff:
            schedule_data.append([e] + [reverse_shift_codes[s] for d in range(7, num_days) for s in all_shifts if solver.Value(works[(e, d, s)]) == 1])
        df_out = pd.DataFrame(schedule_data, columns=['Staff'] + [f'Day_{i}' for i in range(1, target_days + 1)])

        # ==========================================================
        # 🎨 第五階段：輸出加工與翻譯字典 (套用各天專屬班別)
        # ==========================================================
        format_map_base = {
            'N1': 'N1 0000-0800', 'N2': 'N2 0000-0800',
            'E1': 'E1 1600-0000', 'E2': 'E2 1600-0000', 'E3': 'E3 1600-0000'
        }
        format_map_weekday = {
            'D1': 'D1 0730-1600', 'D2': 'DBR 0730-1600', 'D3': 'DBA 0730-1600',
            'D4': 'DCB 0800-1700', 'D5': 'DF 0800-1700', 'D6': 'DF 0800-1700', 'D7': 'DF 0800-1700'
        }
        format_map_sat = {
            'D1': 'D6 0730-1600', 'D2': 'DBR 0800-1600', 'D3': 'DBA 0800-1600',
            'D4': 'D4 0730-1600', 'D5': 'D8 0730-1600', 'D6': 'D9 0800-1600'
        }
        format_map_sun = {
            'D1': 'D4 0800-1600', 'D2': 'DBR 0800-1600', 'D3': 'DBA 0800-1600'
        }

        for col in df_out.columns[1:]:
            day_idx = int(col.split('_')[1])
            d = 6 + day_idx
            is_sat = (d % 7 == 5)
            is_sun = (d % 7 == 6)

            # 依欄位找出排定班別的 index 並洗牌
            idx_N, idx_E, idx_D = df_out[df_out[col]=='N'].index.tolist(), df_out[df_out[col]=='E'].index.tolist(), df_out[df_out[col]=='D'].index.tolist()
            p_N, p_E, p_D = [f'N{i}' for i in range(1, len(idx_N)+1)], [f'E{i}' for i in range(1, len(idx_E)+1)], [f'D{i}' for i in range(1, len(idx_D)+1)]
            random.shuffle(p_N); random.shuffle(p_E); random.shuffle(p_D)

            # 套用基礎與專屬字典
            for i, idx in enumerate(idx_N): df_out.at[idx, col] = format_map_base.get(p_N[i], p_N[i])
            for i, idx in enumerate(idx_E): df_out.at[idx, col] = format_map_base.get(p_E[i], p_E[i])
            for i, idx in enumerate(idx_D):
                code = p_D[i]
                if is_sat: df_out.at[idx, col] = format_map_sat.get(code, 'DF 0800-1700')
                elif is_sun: df_out.at[idx, col] = format_map_sun.get(code, 'DF 0800-1700')
                else:
                    if int(code.replace('D','')) > 7: df_out.at[idx, col] = 'DF 0800-1700'
                    else: df_out.at[idx, col] = format_map_weekday.get(code, code)

        # 套用一例一休轉換機制 (每週重新結算 X1, X2 為 休, 例)
        x_format_map = {1: '休', 2: '例'}
        for idx in df_out.index:
            for w in range(weeks):
                c_X = 1
                for col in [f'Day_{d}' for d in range(w*7 + 1, w*7 + 8)]:
                    if df_out.at[idx, col] == 'X':
                        df_out.at[idx, col] = x_format_map.get(c_X, f'X{c_X}')
                        c_X += 1

        df_out.to_excel(OUTPUT_FILE, index=False)
        print(f"📁 神仙班表已產出：【{OUTPUT_FILE}】！")
    else:
        print("\n❌ 依然排不出來！請檢查人力需求是否過高，或特休畫得太密集。")

# =====================================================================
# 🚀 啟動區域
# =====================================================================
if not os.path.exists(INPUT_FILE):
    create_excel_template()
else:
    run_schedule_optimizer()
