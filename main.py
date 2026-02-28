"""
シフト表自動作成アプリ v3.0
修正内容:
  - パート職員に有給を自動割り当てしない（指定がある場合のみ）
  - パート職員の備考による勤務体系制御（早出のみ等）
  - 固定公休（曜日指定）の対応
  - 週単位勤務日数の柔軟な管理（等式→上下限）
  - スタンドアロン/クラウド両対応
  - ソルバータイムアウト延長(300秒)
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import shutil, os, uuid, re
from ortools.sat.python import cp_model
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

app = FastAPI(title="シフト表自動作成アプリ v3.0")
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

WORK_SHIFTS = ["早", "遅", "夜", "日"]
REST_SHIFTS  = ["×", "有"]
ALL_SHIFTS   = WORK_SHIFTS + REST_SHIFTS

PINK_FILL   = PatternFill("solid", fgColor="FFB6C1")
GREEN_FILL  = PatternFill("solid", fgColor="90EE90")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF99")
GRAY_FILL   = PatternFill("solid", fgColor="D3D3D3")

# 曜日名→weekday()番号
WEEKDAY_MAP = {
    "月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6,
    "月曜": 0, "火曜": 1, "水曜": 2, "木曜": 3, "金曜": 4, "土曜": 5, "日曜": 6,
}


# ============================
# Settings 読み込み
# ============================
def load_settings(df):
    start, end = None, None
    holidays = {}

    header_row = None
    for i in range(len(df)):
        v = str(df.iloc[i, 0]).strip()
        if "期間" in v and "開始" in v:
            header_row = i
            break
    if header_row is None:
        raise Exception("Settingsシートに期間ヘッダーが見つかりません")

    for j in range(header_row + 1, len(df)):
        s = pd.to_datetime(df.iloc[j, 0], errors="coerce")
        e = pd.to_datetime(df.iloc[j, 1], errors="coerce")
        c = str(df.iloc[j, 2]).strip()
        n_str = str(df.iloc[j, 3]).strip()
        if pd.isna(s) and pd.isna(e) and c in ["nan", "None", ""]:
            continue
        if pd.notna(s):
            start = s if start is None else min(start, s)
        if pd.notna(e):
            end = e if end is None else max(end, e)
        m = re.search(r"\d+", n_str)
        if m and c not in ["nan", "None", ""]:
            num = int(m.group())
            if "40" in c:
                holidays["40h"] = holidays.get("40h", 0) + num
            elif "32" in c:
                holidays["32h"] = holidays.get("32h", 0) + num
            elif "パート" in c:
                holidays["パート"] = holidays.get("パート", 0) + num

    if start is None or end is None:
        raise Exception("期間が取得できませんでした")

    holidays.setdefault("40h", 9)
    holidays.setdefault("32h", 8)
    # パートは週単位で管理するため公休数の下限を0にしておく
    holidays.setdefault("パート", 0)

    days = []
    d = start
    while d <= end:
        days.append(d)
        d += timedelta(days=1)
    return days, holidays


# ============================
# 希望シフト 読み込み
# ============================
def load_requests(df, days, staff_list, part_staff=None):
    """
    part_staff: パート職員リスト。
    パート職員の「有給」は、Shift_Requestsで明示的に指定された場合のみ読み込む。
    """
    if part_staff is None:
        part_staff = []
    requests = {}

    header_row = None
    for i in range(len(df)):
        if str(df.iloc[i, 0]).strip() == "職員名":
            header_row = i
            break
    if header_row is None:
        return requests

    col_to_date = {}
    for j in range(1, len(df.columns)):
        d = pd.to_datetime(df.iloc[header_row, j], errors="coerce")
        if pd.notna(d):
            col_to_date[j] = d.to_pydatetime().replace(
                tzinfo=None, hour=0, minute=0, second=0, microsecond=0)

    data_start = header_row + 2
    for i in range(data_start, len(df)):
        name = str(df.iloc[i, 0]).strip()
        if name in ["nan", "None", "", "0"] or name not in staff_list:
            continue

        requests[name] = {}
        is_part = (name in part_staff)

        for j, date in col_to_date.items():
            raw = str(df.iloc[i, j]).strip()
            if raw in ["nan", "None", "", "0"]:
                continue

            if "×" in raw or "休み" in raw:
                requests[name][date] = ("×", "希望")
            elif "有給" in raw or raw == "有":
                if is_part:
                    # パート職員: 明示的な有給指定 → 固定
                    requests[name][date] = ("有", "指定")
                else:
                    requests[name][date] = ("有", "希望")
            elif "夜勤" in raw or raw == "夜":
                requests[name][date] = ("夜", "指定")
            elif "早出" in raw or raw == "早":
                requests[name][date] = ("早", "指定")
            elif "遅出" in raw or raw == "遅":
                requests[name][date] = ("遅", "指定")
            elif "日勤" in raw or raw == "日":
                requests[name][date] = ("日", "指定")

    return requests


# ============================
# 前月実績 読み込み
# ============================
def load_prev_month(df, staff_list):
    prev = {}
    header_row = None
    for i in range(len(df)):
        if str(df.iloc[i, 0]).strip() == "職員名":
            header_row = i
            break
    if header_row is None:
        return prev

    date_cols = []
    for j in range(1, len(df.columns)):
        d = pd.to_datetime(df.iloc[header_row, j], errors="coerce")
        if pd.notna(d):
            date_cols.append(j)

    for i in range(header_row + 1, len(df)):
        name = str(df.iloc[i, 0]).strip()
        if name in ["nan", "None", "", "0"] or name not in staff_list:
            continue
        seq = []
        for j in date_cols:
            raw = str(df.iloc[i, j]).strip()
            if "夜勤" in raw or raw == "夜":
                seq.append("夜")
            elif "早出" in raw or raw == "早":
                seq.append("早")
            elif "遅出" in raw or raw == "遅":
                seq.append("遅")
            elif "日勤" in raw or raw == "日":
                seq.append("日")
            else:
                seq.append("×")
        prev[name] = seq
    return prev


# ============================
# 前月末の連勤カウント
# ============================
def count_trailing_consec(shift_seq):
    count = 0
    for s in reversed(shift_seq):
        if s in ["早", "遅", "夜", "日", "有"]:
            count += 1
        else:
            break
    return count


# ============================
# メインシフト生成
# ============================
def generate_shift(file_path):
    xls = pd.ExcelFile(file_path)
    staff_df    = xls.parse("Staff_Master",   header=None)
    settings_df = xls.parse("Settings",       header=None)
    request_df  = xls.parse("Shift_Requests", header=None)
    prev_df     = xls.parse("Prev_Month",     header=None)

    # Staff_Master ヘッダー行探索
    for i in range(len(staff_df)):
        if str(staff_df.iloc[i, 0]).strip() == "職員名":
            staff_df.columns = staff_df.iloc[i]
            staff_df = staff_df.iloc[i+1:].reset_index(drop=True)
            break

    staff_df = staff_df[staff_df["職員名"].notna()].copy()
    staff_df = staff_df[~staff_df["職員名"].astype(str).isin(["nan","0",""])].copy()
    staff_df["職員名"]    = staff_df["職員名"].astype(str).str.strip()
    staff_df["夜勤最少数"] = pd.to_numeric(staff_df.get("夜勤最少数", pd.Series()), errors="coerce").fillna(0).astype(int)
    staff_df["夜勤最高数"] = pd.to_numeric(staff_df.get("夜勤最高数", pd.Series()), errors="coerce").fillna(0).astype(int)

    staff    = staff_df["職員名"].tolist()
    unit_map = dict(zip(staff_df["職員名"], staff_df["ユニット"].astype(str).str.strip()))
    cont_map = dict(zip(staff_df["職員名"], staff_df["契約区分"].astype(str).str.strip()))
    role_map = dict(zip(staff_df["職員名"], staff_df["役職"].astype(str).str.strip()))
    nmin_map = dict(zip(staff_df["職員名"], staff_df["夜勤最少数"]))
    nmax_map = dict(zip(staff_df["職員名"], staff_df["夜勤最高数"]))

    # 備考列
    note_col = None
    for col_name in staff_df.columns:
        if "備考" in str(col_name):
            note_col = col_name
            break
    if note_col is not None:
        note_map = dict(zip(staff_df["職員名"], staff_df[note_col].astype(str).str.strip()))
    else:
        note_map = {s: "" for s in staff}

    # 固定公休列（固定公休 or 固定休日）
    fixed_hol_col = None
    for col_name in staff_df.columns:
        if "固定" in str(col_name) and ("公休" in str(col_name) or "休" in str(col_name)):
            fixed_hol_col = col_name
            break
    fixed_holiday_map = {}  # name -> list of weekday numbers
    if fixed_hol_col is not None:
        for _, row in staff_df.iterrows():
            name = row["職員名"]
            val  = str(row[fixed_hol_col]).strip()
            if val in ["nan", "None", "", "0", "-"]:
                continue
            wdays = []
            for token in re.split(r"[,、・\s]+", val):
                token = token.strip()
                if token in WEEKDAY_MAP:
                    wdays.append(WEEKDAY_MAP[token])
            if wdays:
                fixed_holiday_map[name] = wdays

    # パート職員リスト
    part_staff = [s for s in staff if cont_map[s] == "パート"]

    # 設定・希望・前月読み込み
    days, holiday_limits = load_settings(settings_df)
    N = len(days)
    requests   = load_requests(request_df, days, staff, part_staff=part_staff)
    prev_month = load_prev_month(prev_df, staff)

    def to_naive(d):
        if hasattr(d, 'to_pydatetime'):
            return d.to_pydatetime().replace(tzinfo=None, hour=0, minute=0, second=0, microsecond=0)
        return datetime(d.year, d.month, d.day)

    days_norm = [to_naive(d) for d in days]

    # ============================
    # 備考解析
    # ============================
    allowed_shifts_map = {}  # s -> set of allowed work shifts (None = 制限なし)
    weekly_work_days   = {}  # s -> 週勤務日数

    for s in staff:
        note = note_map.get(s, "")
        allowed = None

        if "早出のみ" in note:
            allowed = {"早"}
        elif "遅出のみ" in note:
            allowed = {"遅"}
        elif "夜勤なし" in note or "夜勤禁止" in note:
            allowed = {"早", "遅", "日"}

        if allowed is not None:
            allowed_shifts_map[s] = allowed

        m = re.search(r"週(\d+)日", note)
        if m:
            weekly_work_days[s] = int(m.group(1))

    # ============================
    # 週グループ（日曜始まり）
    # ============================
    week_groups = defaultdict(list)
    for d_idx, dn in enumerate(days_norm):
        sun_offset = (dn.weekday() + 1) % 7
        week_sun   = dn - timedelta(days=sun_offset)
        week_key   = week_sun.strftime("%Y-%m-%d")
        week_groups[week_key].append(d_idx)
    sorted_week_keys = sorted(week_groups.keys())

    # A・B 職員（両ユニット掛け持ち）
    ab_staff = [s for s in staff if unit_map[s] == "A・B"]

    # ========== CP-SAT モデル ==========
    model = cp_model.CpModel()

    x = {}
    for s in staff:
        for d in range(N):
            for sh in ALL_SHIFTS:
                x[s, d, sh] = model.NewBoolVar(f"x_{s}_{d}_{sh}")

    # A・B 職員のユニット割り当て変数
    uea = {}; ueb = {}; ula = {}; ulb = {}
    for s in ab_staff:
        for d in range(N):
            uea[s, d] = model.NewBoolVar(f"uea_{s}_{d}")
            ueb[s, d] = model.NewBoolVar(f"ueb_{s}_{d}")
            ula[s, d] = model.NewBoolVar(f"ula_{s}_{d}")
            ulb[s, d] = model.NewBoolVar(f"ulb_{s}_{d}")
            model.Add(uea[s, d] + ueb[s, d] == x[s, d, "早"])
            model.Add(ula[s, d] + ulb[s, d] == x[s, d, "遅"])

    # ---------- 制約1: 1日1シフト ----------
    for s in staff:
        for d in range(N):
            model.AddExactlyOne(x[s, d, sh] for sh in ALL_SHIFTS)

    # ---------- 制約2: 希望シフト固定 ----------
    for s in staff:
        if s not in requests:
            continue
        for date_obj, (sh_type, _) in requests[s].items():
            for d, dn in enumerate(days_norm):
                if dn == date_obj:
                    if sh_type in ALL_SHIFTS:
                        model.Add(x[s, d, sh_type] == 1)
                    break

    # ---------- 制約3: 前月最終日が夜勤 → 1日目は× ----------
    for s in staff:
        seq = prev_month.get(s, [])
        if seq and seq[-1] == "夜":
            model.Add(x[s, 0, "×"] == 1)

    # ---------- 制約4: 固定公休（曜日指定）----------
    for s, wdays in fixed_holiday_map.items():
        for d_idx, dn in enumerate(days_norm):
            if dn.weekday() in wdays:
                # Shift_Requestsで上書き指定がある日は除外
                date_obj = dn
                req = requests.get(s, {}).get(date_obj)
                if req and req[1] == "指定":
                    continue
                model.Add(x[s, d_idx, "×"] == 1)

    # ---------- 制約5: 毎日の必須人数（A・B職員は一方のみカウント） ----------
    for d in range(N):
        # A早出 = Aユニット専属の早出 + AB職員のA側早出
        a_early = [x[s, d, "早"] for s in staff if unit_map[s] == "A"]
        a_early += [uea[s, d] for s in ab_staff]
        model.Add(sum(a_early) == 1)

        a_late = [x[s, d, "遅"] for s in staff if unit_map[s] == "A"]
        a_late += [ula[s, d] for s in ab_staff]
        model.Add(sum(a_late) == 1)

        b_early = [x[s, d, "早"] for s in staff if unit_map[s] == "B"]
        b_early += [ueb[s, d] for s in ab_staff]
        model.Add(sum(b_early) == 1)

        b_late = [x[s, d, "遅"] for s in staff if unit_map[s] == "B"]
        b_late += [ulb[s, d] for s in ab_staff]
        model.Add(sum(b_late) == 1)

        model.Add(sum(x[s, d, "夜"] for s in staff) == 1)

    # ---------- 制約6: 夜勤回数 ----------
    for s in staff:
        night_total = sum(x[s, d, "夜"] for d in range(N))
        model.Add(night_total >= nmin_map[s])
        model.Add(night_total <= nmax_map[s])

    # ---------- 制約7: 夜勤 → 翌日× ----------
    for s in staff:
        for d in range(N - 1):
            model.Add(x[s, d+1, "×"] == 1).OnlyEnforceIf(x[s, d, "夜"])

    # ---------- 制約8: 遅 → 翌早 禁止 ----------
    for s in staff:
        for d in range(N - 1):
            model.Add(x[s, d, "遅"] + x[s, d+1, "早"] <= 1)

    # ---------- 制約9: 希望休の前日に夜勤を入れない ----------
    for s in staff:
        if s not in requests:
            continue
        for date_obj, (sh_type, req_type) in requests[s].items():
            if req_type == "希望" and sh_type in ["×", "有"]:
                for d, dn in enumerate(days_norm):
                    if dn == date_obj:
                        if d > 0:
                            model.Add(x[s, d-1, "夜"] == 0)
                        break

    # ---------- 制約10: 連勤制限 ----------
    for s in staff:
        max_c  = 5 if cont_map[s] == "40h" else 4
        prev_c = count_trailing_consec(prev_month.get(s, []))
        remain = max(0, max_c - prev_c)

        if prev_c > 0 and remain < max_c:
            # 月頭のwindow
            for w in range(1, min(remain + 2, N + 1)):
                if w > remain:
                    model.Add(
                        sum(x[s, d2, sh2]
                            for d2 in range(w)
                            for sh2 in ["早","遅","夜","有","日"])
                        <= remain
                    )
                    break

        for st in range(N - max_c):
            model.Add(
                sum(x[s, d2, sh2]
                    for d2 in range(st, st + max_c + 1)
                    for sh2 in ["早","遅","夜","有","日"])
                <= max_c
            )

    # ---------- 制約11: 公休数確保 ----------
    for s in staff:
        min_hol = holiday_limits.get(cont_map[s], 8)
        if min_hol > 0:
            model.Add(sum(x[s, d, "×"] for d in range(N)) >= min_hol)

    # ---------- 制約12: 備考による勤務制限 ----------
    for s in staff:
        allowed = allowed_shifts_map.get(s)
        if allowed is None:
            continue
        forbidden = set(WORK_SHIFTS) - allowed
        for d in range(N):
            for sh in forbidden:
                date_obj = days_norm[d]
                req = requests.get(s, {}).get(date_obj)
                if req and req[0] == sh and req[1] == "指定":
                    continue
                model.Add(x[s, d, sh] == 0)

    # ---------- 制約13: パート職員に有給を自動割り当てしない ----------
    for s in part_staff:
        for d in range(N):
            date_obj = days_norm[d]
            req = requests.get(s, {}).get(date_obj)
            # Shift_Requestsで「有」と明示指定されていない限り有給禁止
            if req and req[0] == "有" and req[1] == "指定":
                pass  # 固定済み
            else:
                model.Add(x[s, d, "有"] == 0)

    # ---------- 制約14: パート職員の週単位勤務日数 ----------
    # 完全週: target-1 ≤ 勤務日数 ≤ target （± 1 日の余裕）
    # 不完全週: 0 ≤ 勤務日数 ≤ target（上限のみ）
    for s in staff:
        if s not in weekly_work_days:
            continue
        target = weekly_work_days[s]
        for week_key in sorted_week_keys:
            didx = week_groups[week_key]
            work_vars = [x[s, d, sh]
                         for d in didx
                         for sh in ["早","遅","夜","有","日"]]
            if len(didx) == 7:
                # 完全週: target ± 1
                model.Add(sum(work_vars) >= max(0, target - 1))
                model.Add(sum(work_vars) <= target)
            else:
                # 不完全週（月初/月末）: 比例配分の上限
                partial_max = round(target * len(didx) / 7 + 0.5)
                model.Add(sum(work_vars) <= partial_max)

    # ========== 目的関数: 早・遅の平準化（リーダー以外） ==========
    non_leader = [s for s in staff if role_map.get(s) != "リーダー"]
    if len(non_leader) >= 2:
        early_vars = []
        late_vars  = []
        for s in non_leader:
            ev = model.NewIntVar(0, N, f"e_{s}")
            lv = model.NewIntVar(0, N, f"l_{s}")
            model.Add(ev == sum(x[s, d, "早"] for d in range(N)))
            model.Add(lv == sum(x[s, d, "遅"] for d in range(N)))
            early_vars.append(ev)
            late_vars.append(lv)

        max_e = model.NewIntVar(0, N, "max_e"); min_e = model.NewIntVar(0, N, "min_e")
        max_l = model.NewIntVar(0, N, "max_l"); min_l = model.NewIntVar(0, N, "min_l")
        model.AddMaxEquality(max_e, early_vars); model.AddMinEquality(min_e, early_vars)
        model.AddMaxEquality(max_l, late_vars);  model.AddMinEquality(min_l, late_vars)
        diff_e = model.NewIntVar(0, N, "diff_e"); model.Add(diff_e == max_e - min_e)
        diff_l = model.NewIntVar(0, N, "diff_l"); model.Add(diff_l == max_l - min_l)
        model.Minimize(diff_e + diff_l)

    # ========== ソルバー ==========
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 300  # 5分
    solver.parameters.num_search_workers  = 8    # 並列化
    status = solver.Solve(model)

    if status not in (cp_model.FEASIBLE, cp_model.OPTIMAL):
        raise Exception(
            "条件を満たすシフト表が見つかりませんでした。\n"
            "希望シフト・夜勤回数・公休数の設定を見直してください。"
        )

    # ========== 結果組み立て ==========
    result = {}
    for s in staff:
        result[s] = {}
        for d in range(N):
            for sh in ALL_SHIFTS:
                if solver.Value(x[s, d, sh]) == 1:
                    result[s][d] = sh
                    break

    # A・B職員のユニット割り当て結果
    ab_unit_result = {}
    for s in ab_staff:
        ab_unit_result[s] = {}
        for d in range(N):
            sh = result[s][d]
            if sh == "早":
                ab_unit_result[s][d] = "A" if solver.Value(uea[s, d]) == 1 else "B"
            elif sh == "遅":
                ab_unit_result[s][d] = "A" if solver.Value(ula[s, d]) == 1 else "B"
            else:
                ab_unit_result[s][d] = None

    return result, staff, unit_map, cont_map, role_map, days_norm, requests, ab_unit_result


# ============================
# Excelへの書き出し
# ============================
def write_shift_result(result, staff, unit_map, cont_map, role_map,
                       days_norm, requests, ab_unit_result,
                       input_path, output_path):
    shutil.copy(input_path, output_path)
    wb = load_workbook(output_path)

    if "shift_result" in wb.sheetnames:
        del wb["shift_result"]
    ws = wb.create_sheet("shift_result")

    N = len(days_norm)
    weekday_ja = ["月","火","水","木","金","土","日"]

    DATE_START_COL = 3
    SUMMARY_COL    = DATE_START_COL + N
    SUMMARY_HDRS   = ["早出", "遅出", "日勤", "夜勤", "公休"]
    STAFF_START_ROW  = 4
    SUMMARY_ROW_BASE = STAFF_START_ROW + len(staff) + 1

    # ===== ヘッダー =====
    ws.cell(1, 1, "作成月")
    ws.cell(1, 2, days_norm[0].strftime("%Y年%m月"))
    ws.cell(2, 2, "曜日")
    ws.cell(3, 1, "ユニット")
    ws.cell(3, 2, "職員名")

    for i, d in enumerate(days_norm):
        col = DATE_START_COL + i
        ws.cell(1, col, d.day).alignment = Alignment(horizontal="center")
        wd = weekday_ja[d.weekday()]
        cell = ws.cell(2, col, wd)
        cell.alignment = Alignment(horizontal="center")
        if d.weekday() == 5:
            cell.fill = PatternFill("solid", fgColor="CCE5FF")
        elif d.weekday() == 6:
            cell.fill = PatternFill("solid", fgColor="FFCCCC")

    for k, h in enumerate(SUMMARY_HDRS):
        c = ws.cell(3, SUMMARY_COL + k, h)
        c.fill = YELLOW_FILL
        c.alignment = Alignment(horizontal="center")

    ws.cell(3, 1).fill = YELLOW_FILL
    ws.cell(3, 2).fill = YELLOW_FILL

    # ===== 職員データ =====
    def unit_order(s):
        u = unit_map[s]
        if u == "A":    return 0
        if u == "A・B": return 1
        return 2

    sorted_staff = sorted(staff, key=unit_order)

    for idx, s in enumerate(sorted_staff):
        row = STAFF_START_ROW + idx
        ws.cell(row, 1, unit_map[s]).alignment = Alignment(horizontal="center")
        ws.cell(row, 2, s).alignment = Alignment(horizontal="center")

        for d in range(N):
            col  = DATE_START_COL + d
            sh   = result[s][d]
            cell = ws.cell(row, col, sh)
            cell.alignment = Alignment(horizontal="center")

            date_obj = days_norm[d]
            if s in requests and date_obj in requests[s]:
                _, req_type = requests[s][date_obj]
                if req_type == "希望":
                    cell.fill = PINK_FILL
                elif req_type == "指定":
                    cell.fill = GREEN_FILL

        ds  = get_column_letter(DATE_START_COL)
        de  = get_column_letter(DATE_START_COL + N - 1)
        rng = f"{ds}{row}:{de}{row}"
        ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"早")')
        ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"遅")')
        ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"日")')
        ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')
        ws.cell(row, SUMMARY_COL + 4, f'=COUNTIF({rng},"×")')

    # ===== 日別集計行 =====
    ab_staff_local = [s for s in staff if unit_map[s] == "A・B"]
    label_names = ["A早出", "B早出", "A遅出", "B遅出", "夜勤"]
    for k, lbl in enumerate(label_names):
        r = SUMMARY_ROW_BASE + k
        ws.cell(r, 2, lbl).fill = GRAY_FILL
        ws.cell(r, 2).alignment = Alignment(horizontal="center")

    for i in range(N):
        d   = i
        col = DATE_START_COL + i

        cnt_a_early = sum(1 for s in staff if unit_map[s]=="A" and result[s][d]=="早")
        cnt_a_early += sum(1 for s in ab_staff_local
                           if ab_unit_result.get(s,{}).get(d)=="A" and result[s][d]=="早")
        cnt_b_early = sum(1 for s in staff if unit_map[s]=="B" and result[s][d]=="早")
        cnt_b_early += sum(1 for s in ab_staff_local
                           if ab_unit_result.get(s,{}).get(d)=="B" and result[s][d]=="早")
        cnt_a_late  = sum(1 for s in staff if unit_map[s]=="A" and result[s][d]=="遅")
        cnt_a_late  += sum(1 for s in ab_staff_local
                           if ab_unit_result.get(s,{}).get(d)=="A" and result[s][d]=="遅")
        cnt_b_late  = sum(1 for s in staff if unit_map[s]=="B" and result[s][d]=="遅")
        cnt_b_late  += sum(1 for s in ab_staff_local
                           if ab_unit_result.get(s,{}).get(d)=="B" and result[s][d]=="遅")
        cnt_night   = sum(1 for s in staff if result[s][d]=="夜")

        for k, v in enumerate([cnt_a_early, cnt_b_early, cnt_a_late, cnt_b_late, cnt_night]):
            ws.cell(SUMMARY_ROW_BASE + k, col, v).alignment = Alignment(horizontal="center")

    # 列幅
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 8
    for i in range(N):
        ws.column_dimensions[get_column_letter(DATE_START_COL + i)].width = 4
    for k in range(len(SUMMARY_HDRS)):
        ws.column_dimensions[get_column_letter(SUMMARY_COL + k)].width = 6

    wb.save(output_path)


# ============================
# Web UI (HTML)
# ============================
HTML_CONTENT = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>シフト表自動作成アプリ v3.0</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;display:flex;justify-content:center;align-items:flex-start;padding:30px 20px}
.card{background:#fff;padding:40px;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-width:900px;width:100%}
h1{color:#667eea;font-size:1.9em;text-align:center;margin-bottom:6px}
.ver{text-align:center;color:#764ba2;font-weight:bold;margin-bottom:4px;font-size:.9em}
.sub{text-align:center;color:#888;margin-bottom:24px;font-size:.85em}
.sec-title{font-weight:bold;color:#333;margin-bottom:10px;font-size:1em;border-left:4px solid #667eea;padding-left:10px;margin-top:16px}
.rules{background:#f8f9fa;padding:16px 20px;border-radius:10px;margin-bottom:14px}
.rules ul{list-style:none}
.rules li{padding:5px 0;border-bottom:1px solid #eee;font-size:.88em;color:#555}
.rules li:last-child{border-bottom:none}
.note{background:#fff8e1;border-left:4px solid #ffc107;padding:12px 16px;border-radius:5px;margin-bottom:18px;font-size:.87em;color:#555}
.drop{border:3px dashed #667eea;border-radius:12px;padding:36px;text-align:center;cursor:pointer;transition:.3s}
.drop:hover,.drop.over{background:#f0f4ff;border-color:#764ba2}
input[type=file]{display:none}
.pick-btn{background:#667eea;color:#fff;padding:9px 26px;border:none;border-radius:22px;cursor:pointer;font-size:.93em;margin-top:10px;display:inline-block}
.pick-btn:hover{background:#764ba2}
.fname{margin-top:10px;color:#555;font-weight:bold;font-size:.88em}
.go-btn{width:100%;background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;padding:13px;border:none;border-radius:22px;font-size:1em;cursor:pointer;margin-top:16px;transition:.3s}
.go-btn:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 22px rgba(102,126,234,.4)}
.go-btn:disabled{background:#ccc;cursor:not-allowed}
.spin-wrap{display:none;text-align:center;margin-top:20px}
.spinner{border:4px solid #eee;border-top:4px solid #667eea;border-radius:50%;width:44px;height:44px;animation:spin 1s linear infinite;margin:0 auto 10px}
@keyframes spin{to{transform:rotate(360deg)}}
.progress-msg{color:#667eea;font-size:.9em}
.ok{display:none;background:#d4edda;border:1px solid #c3e6cb;color:#155724;padding:16px;border-radius:10px;margin-top:16px;text-align:center}
.dl-btn{display:inline-block;background:#28a745;color:#fff;padding:10px 28px;text-decoration:none;border-radius:20px;margin-top:10px;font-size:.95em}
.dl-btn:hover{background:#218838}
.err{display:none;background:#f8d7da;border:1px solid #f5c6cb;color:#721c24;padding:14px;border-radius:10px;margin-top:16px;word-break:break-all;white-space:pre-wrap;font-size:.88em}
.legend{display:flex;gap:16px;margin-top:14px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px;font-size:.83em;color:#555}
.swatch{width:16px;height:16px;border-radius:3px;border:1px solid #ccc}
.pink{background:#FFB6C1}.green{background:#90EE90}
</style>
</head>
<body>
<div class="card">
  <h1>📅 シフト表自動作成アプリ</h1>
  <p class="ver">Version 3.0</p>
  <p class="sub">Excelファイルをアップロードするだけで最適なシフト表を自動生成</p>

  <div class="sec-title">🔒 適用される制約条件</div>
  <div class="rules"><ul>
    <li>✅ ユニットA/B：毎日 <strong>早出1・遅出1</strong>（A・B職員はどちらか一方にカウント）</li>
    <li>✅ 夜勤：毎日1名（全体）、個人の <strong>最少〜最高回数</strong> を厳守</li>
    <li>✅ 40h→最大5連勤 / 32h・パート→最大4連勤（前月継続分を考慮）</li>
    <li>✅ 夜勤→翌日必ず×、遅出→翌日早出禁止</li>
    <li>✅ 希望休の <strong>前日に夜勤を入れない</strong></li>
    <li>✅ <strong>パート職員：有給を自動割り当てしない</strong>（Shift_Requestsで指定がある場合のみ）</li>
    <li>✅ Staff_Masterの <strong>備考を厳守</strong>（早出のみ・遅出のみ・夜勤なし等）</li>
    <li>✅ <strong>固定公休</strong>（例：日曜固定）を曜日で指定可能</li>
    <li>✅ パート職員の <strong>週単位勤務日数</strong>（日〜土）を管理</li>
    <li>✅ 希望休→ピンク・勤務指定→緑でExcelに色付け</li>
    <li>✅ 各職員の公休数を確保・リーダー以外の早遅を平準化</li>
  </ul></div>

  <div class="note">
    <strong>📋 必要なシート（5枚）：</strong>
    Staff_Master / Settings / Shift_Requests / Prev_Month / shift_result<br>
    <strong>備考欄の例：</strong>「早出のみ。週4日勤務。」「週5日勤務。夜勤なし。」<br>
    <strong>固定公休欄の例：</strong>「日曜」「土・日」など曜日を記入
  </div>

  <div class="sec-title">📤 ファイルアップロード</div>
  <form id="frm">
    <div class="drop" id="drop">
      <p>📂 ここにExcelファイルをドラッグ＆ドロップ</p>
      <p style="margin:8px 0;color:#aaa">— または —</p>
      <label for="fi" class="pick-btn">ファイルを選択</label>
      <input type="file" id="fi" accept=".xlsx,.xls">
      <div class="fname" id="fname"></div>
    </div>
    <button type="submit" class="go-btn" id="go">▶ シフト表を生成する</button>
  </form>

  <div class="spin-wrap" id="sw">
    <div class="spinner"></div>
    <p class="progress-msg" id="pmsg">生成中… 最大5分かかる場合があります<br>このままお待ちください</p>
  </div>
  <div class="ok" id="ok">
    <p>✅ シフト表の生成が完了しました！</p>
    <a href="#" id="dl" class="dl-btn">📥 Shift_Result.xlsx をダウンロード</a>
  </div>
  <div class="err" id="er"></div>

  <div class="legend">
    <div class="legend-item"><div class="swatch pink"></div>希望休・有給（希望）</div>
    <div class="legend-item"><div class="swatch green"></div>勤務指定（指定）</div>
  </div>
</div>
<script>
const fi=document.getElementById('fi'),fname=document.getElementById('fname'),
      drop=document.getElementById('drop'),frm=document.getElementById('frm'),
      sw=document.getElementById('sw'),ok=document.getElementById('ok'),
      er=document.getElementById('er'),dl=document.getElementById('dl'),
      go=document.getElementById('go'),pmsg=document.getElementById('pmsg');

fi.onchange=()=>{ if(fi.files[0]) fname.textContent='📄 '+fi.files[0].name; };

['dragenter','dragover','dragleave','drop'].forEach(e=>
  drop.addEventListener(e,ev=>{ev.preventDefault();ev.stopPropagation();}));
['dragenter','dragover'].forEach(e=>drop.addEventListener(e,()=>drop.classList.add('over')));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,()=>drop.classList.remove('over')));
drop.addEventListener('drop',e=>{
  const files=e.dataTransfer.files;
  if(files[0]){
    const dt=new DataTransfer(); dt.items.add(files[0]);
    fi.files=dt.files;
    fname.textContent='📄 '+files[0].name;
  }
});
drop.addEventListener('click',()=>fi.click());

let elapsed=0, timer=null;
function startTimer(){
  elapsed=0;
  timer=setInterval(()=>{
    elapsed++;
    pmsg.innerHTML='生成中… <strong>'+elapsed+'秒</strong> 経過<br>最大5分かかる場合があります。このままお待ちください';
  },1000);
}
function stopTimer(){ if(timer){clearInterval(timer);timer=null;} }

frm.onsubmit=async e=>{
  e.preventDefault();
  if(!fi.files[0]){alert('ファイルを選択してください');return;}
  const fd=new FormData(); fd.append('file',fi.files[0]);
  sw.style.display='block'; ok.style.display='none';
  er.style.display='none'; go.disabled=true;
  startTimer();
  try{
    const res=await fetch('/generate-shift',{method:'POST',body:fd});
    stopTimer();
    if(res.ok){
      const blob=await res.blob();
      dl.href=URL.createObjectURL(blob);
      dl.download='Shift_Result.xlsx';
      sw.style.display='none'; ok.style.display='block';
    }else{
      const j=await res.json().catch(()=>({}));
      throw new Error(j.detail||'サーバーエラーが発生しました');
    }
  }catch(ex){
    stopTimer();
    sw.style.display='none';
    er.style.display='block';
    er.textContent='❌ エラー:\\n'+ex.message;
  }finally{ go.disabled=false; }
};
</script>
</body>
</html>"""


# ============================
# FastAPI Routes
# ============================
@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(content=HTML_CONTENT)


@app.get("/health")
async def health():
    return {"status": "ok", "version": "3.0"}


@app.post("/generate-shift")
async def generate(file: UploadFile = File(...)):
    uid      = str(uuid.uuid4())
    in_path  = os.path.join(TEMP_DIR, f"in_{uid}.xlsx")
    out_path = os.path.join(TEMP_DIR, f"out_{uid}.xlsx")
    try:
        with open(in_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        result, staff, unit_map, cont_map, role_map, days_norm, requests, ab_unit_result = \
            generate_shift(in_path)

        write_shift_result(
            result, staff, unit_map, cont_map, role_map,
            days_norm, requests, ab_unit_result,
            in_path, out_path
        )

        return FileResponse(
            out_path,
            filename="Shift_Result.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        try: os.remove(in_path)
        except: pass


# ============================
# スタンドアロン起動
# ============================
if __name__ == "__main__":
    import uvicorn, webbrowser, threading, time

    def open_browser():
        time.sleep(2.0)
        webbrowser.open("http://localhost:8000")

    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")

    # ローカル起動時のみブラウザを自動で開く
    if os.environ.get("AUTO_BROWSER", "1") == "1" and port == 8000:
        threading.Thread(target=open_browser, daemon=True).start()

    print("=" * 50)
    print(" シフト表自動作成アプリ v3.0")
    print(f" http://localhost:{port}")
    print("=" * 50)
    uvicorn.run("main:app", host=host, port=port, reload=False)
