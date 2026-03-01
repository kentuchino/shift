"""
シフト表自動作成アプリ v4.0
新機能:
  - 公休日数をなるべく指定数に近づける（リーダー以外）
  - 連続夜勤（Staff_Masterで○指定の職員のみ緊急時に許可）
  - 勤務間隔：なるべく3〜4日に1回休み（ソフト制約）
  - 同一勤務の連続を避ける（ソフト制約、パート指定除く）
  - 主任：本来の職員だけでは組めない時のみ早出で使用
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

app = FastAPI(title="シフト表自動作成アプリ v4.0")
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

WORK_SHIFTS = ["早", "遅", "夜", "日"]
REST_SHIFTS  = ["×", "有"]
ALL_SHIFTS   = WORK_SHIFTS + REST_SHIFTS

PINK_FILL   = PatternFill("solid", fgColor="FFB6C1")
GREEN_FILL  = PatternFill("solid", fgColor="90EE90")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF99")
GRAY_FILL   = PatternFill("solid", fgColor="D3D3D3")
BLUE_FILL   = PatternFill("solid", fgColor="BDD7EE")   # 主任使用日

WEEKDAY_MAP = {
    "月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6,
    "月曜": 0, "火曜": 1, "水曜": 2, "木曜": 3, "金曜": 4, "土曜": 5, "日曜": 6,
}


# ========================================================
# Settings 読み込み
# ========================================================
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
    holidays.setdefault("パート", 0)

    days = []
    d = start
    while d <= end:
        days.append(d)
        d += timedelta(days=1)
    return days, holidays


# ========================================================
# 希望シフト 読み込み
# ========================================================
def load_requests(df, days, staff_list, part_staff=None):
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
                requests[name][date] = ("有", "指定" if is_part else "希望")
            elif "夜勤" in raw or raw == "夜":
                requests[name][date] = ("夜", "指定")
            elif "早出" in raw or raw == "早":
                requests[name][date] = ("早", "指定")
            elif "遅出" in raw or raw == "遅":
                requests[name][date] = ("遅", "指定")
            elif "日勤" in raw or raw == "日":
                requests[name][date] = ("日", "指定")
    return requests


# ========================================================
# 前月実績 読み込み
# ========================================================
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
            if "夜勤" in raw or raw == "夜":   seq.append("夜")
            elif "早出" in raw or raw == "早": seq.append("早")
            elif "遅出" in raw or raw == "遅": seq.append("遅")
            elif "日勤" in raw or raw == "日": seq.append("日")
            else:                              seq.append("×")
        prev[name] = seq
    return prev


def count_trailing_consec(shift_seq):
    count = 0
    for s in reversed(shift_seq):
        if s in ["早", "遅", "夜", "日", "有"]:
            count += 1
        else:
            break
    return count


# ========================================================
# メインシフト生成
# ========================================================
def generate_shift(file_path):
    xls = pd.ExcelFile(file_path)
    staff_df    = xls.parse("Staff_Master",   header=None)
    settings_df = xls.parse("Settings",       header=None)
    request_df  = xls.parse("Shift_Requests", header=None)
    prev_df     = xls.parse("Prev_Month",     header=None)

    # ── Staff_Master 読み込み ──
    for i in range(len(staff_df)):
        if str(staff_df.iloc[i, 0]).strip() == "職員名":
            staff_df.columns = staff_df.iloc[i]
            staff_df = staff_df.iloc[i+1:].reset_index(drop=True)
            break

    staff_df = staff_df[staff_df["職員名"].notna()].copy()
    staff_df = staff_df[~staff_df["職員名"].astype(str).isin(["nan","0",""])].copy()
    staff_df["職員名"] = staff_df["職員名"].astype(str).str.strip()

    def col_num(name, default=0):
        if name in staff_df.columns:
            return pd.to_numeric(staff_df[name], errors="coerce").fillna(default).astype(int)
        return pd.Series([default]*len(staff_df))

    staff_df["夜勤最少数"] = col_num("夜勤最少数", 0)
    staff_df["夜勤最高数"] = col_num("夜勤最高数", 0)

    all_staff_names = staff_df["職員名"].tolist()

    def get_map(col, default=""):
        if col in staff_df.columns:
            return dict(zip(staff_df["職員名"], staff_df[col].astype(str).str.strip()))
        return {s: default for s in all_staff_names}

    unit_map  = get_map("ユニット")
    cont_map  = get_map("契約区分")
    role_map  = get_map("役職")
    nmin_map  = dict(zip(staff_df["職員名"], staff_df["夜勤最少数"]))
    nmax_map  = dict(zip(staff_df["職員名"], staff_df["夜勤最高数"]))
    note_map  = get_map("備考")
    # 連続夜勤: ○ の職員のみ許可
    consec_night_map = get_map("連続夜勤")   # "○" or "×"

    # 固定公休
    fixed_holiday_map = {}
    fhcol = next((c for c in staff_df.columns if "固定" in str(c) and "休" in str(c)), None)
    if fhcol:
        for _, row in staff_df.iterrows():
            val = str(row[fhcol]).strip()
            if val in ["nan","None","","-","0"]:
                continue
            wdays = [WEEKDAY_MAP[t.strip()] for t in re.split(r"[,、・\s]+", val)
                     if t.strip() in WEEKDAY_MAP]
            if wdays:
                fixed_holiday_map[row["職員名"]] = wdays

    # 主任の識別
    SHUUNIN_NAME = "主任"
    shuunin_list = [s for s in all_staff_names
                    if role_map.get(s, "") in ("総合","主任") and
                       unit_map.get(s, "") in ("nan","","NaN")]
    # ユニット欄がnull/nanの場合を主任判定
    shuunin_list = [s for s in all_staff_names
                    if str(unit_map.get(s, "")).lower() in ("nan", "", "none")]

    # 通常スタッフ（主任除く）
    staff = [s for s in all_staff_names if s not in shuunin_list]
    part_staff = [s for s in staff if cont_map[s] == "パート"]

    # 設定・希望・前月
    days, holiday_limits = load_settings(settings_df)
    N = len(days)
    all_names_for_req = all_staff_names  # 主任も希望シフト対象
    requests   = load_requests(request_df, days, all_names_for_req, part_staff=part_staff)
    prev_month = load_prev_month(prev_df, all_names_for_req)

    def to_naive(d):
        if hasattr(d, 'to_pydatetime'):
            return d.to_pydatetime().replace(tzinfo=None, hour=0, minute=0, second=0, microsecond=0)
        return datetime(d.year, d.month, d.day)
    days_norm = [to_naive(d) for d in days]

    # ── 備考解析 ──
    allowed_shifts_map = {}
    weekly_work_days   = {}
    # パート職員で勤務指定がある（Shift_Requestsに指定あり）= 同一勤務連続ペナルティを除外
    part_with_fixed = set()

    for s in all_staff_names:
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

    for s in part_staff:
        req_s = requests.get(s, {})
        designated = sum(1 for v in req_s.values() if v[1] == "指定" and v[0] in WORK_SHIFTS)
        if designated >= 3:
            part_with_fixed.add(s)

    # 週グループ
    week_groups = defaultdict(list)
    for d_idx, dn in enumerate(days_norm):
        sun_offset = (dn.weekday() + 1) % 7
        week_sun   = dn - timedelta(days=sun_offset)
        week_groups[week_sun.strftime("%Y-%m-%d")].append(d_idx)
    sorted_week_keys = sorted(week_groups.keys())

    # A・B 兼務職員
    ab_staff = [s for s in staff if unit_map.get(s, "") == "A・B"]

    # ========================================================
    # CP-SAT モデル
    # ========================================================
    model = cp_model.CpModel()

    # 通常スタッフ変数
    x = {}
    for s in staff:
        for d in range(N):
            for sh in ALL_SHIFTS:
                x[s, d, sh] = model.NewBoolVar(f"x_{s}_{d}_{sh}")

    # 主任変数
    xs = {}  # xs[shuunin_name, d, sh]
    for s in shuunin_list:
        for d in range(N):
            for sh in ALL_SHIFTS:
                xs[s, d, sh] = model.NewBoolVar(f"xs_{s}_{d}_{sh}")

    # A・B 兼務ユニット割り当て変数
    uea = {}; ueb = {}; ula = {}; ulb = {}
    for s in ab_staff:
        for d in range(N):
            uea[s,d] = model.NewBoolVar(f"uea_{s}_{d}")
            ueb[s,d] = model.NewBoolVar(f"ueb_{s}_{d}")
            ula[s,d] = model.NewBoolVar(f"ula_{s}_{d}")
            ulb[s,d] = model.NewBoolVar(f"ulb_{s}_{d}")
            model.Add(uea[s,d] + ueb[s,d] == x[s,d,"早"])
            model.Add(ula[s,d] + ulb[s,d] == x[s,d,"遅"])

    # 主任ユニット補完変数
    shuunin_use_a = {}; shuunin_use_b = {}
    for s in shuunin_list:
        for d in range(N):
            shuunin_use_a[s,d] = model.NewBoolVar(f"sh_ua_{s}_{d}")
            shuunin_use_b[s,d] = model.NewBoolVar(f"sh_ub_{s}_{d}")
            # 主任が早出の日のみ補完可
            model.Add(shuunin_use_a[s,d] + shuunin_use_b[s,d] <= xs[s,d,"早"])
            # 主任は同日にA・Bどちらか一方のみ
            model.Add(shuunin_use_a[s,d] + shuunin_use_b[s,d] <= 1)

    # ── 制約1: 1日1シフト ──
    for s in staff:
        for d in range(N):
            model.AddExactlyOne(x[s,d,sh] for sh in ALL_SHIFTS)
    for s in shuunin_list:
        for d in range(N):
            model.AddExactlyOne(xs[s,d,sh] for sh in ALL_SHIFTS)

    # ── 制約2: 希望・指定シフト固定 ──
    def fix_requests(var_dict, s_list):
        for s in s_list:
            if s not in requests:
                continue
            for date_obj, (sh_type, _) in requests[s].items():
                for d, dn in enumerate(days_norm):
                    if dn == date_obj and sh_type in ALL_SHIFTS:
                        model.Add(var_dict[s,d,sh_type] == 1)
                        break
    fix_requests(x, staff)
    fix_requests(xs, shuunin_list)

    # ── 制約3: 前月最終日が夜勤 → 1日目は× ──
    for s in staff:
        if prev_month.get(s, []) and prev_month[s][-1] == "夜":
            model.Add(x[s,0,"×"] == 1)
    for s in shuunin_list:
        if prev_month.get(s, []) and prev_month[s][-1] == "夜":
            model.Add(xs[s,0,"×"] == 1)

    # ── 制約4: 固定公休（曜日指定）──
    for s, wdays in fixed_holiday_map.items():
        var_dict = xs if s in shuunin_list else x
        for d_idx, dn in enumerate(days_norm):
            if dn.weekday() in wdays:
                req = requests.get(s, {}).get(dn)
                if req and req[1] == "指定":
                    continue
                model.Add(var_dict[s,d_idx,"×"] == 1)

    # ── 制約5: 毎日の必須人数 ──
    # 主任はどうしても組めない場合のみ補完（ペナルティで制御）
    for d in range(N):
        # A早出
        a_e = [x[s,d,"早"] for s in staff if unit_map.get(s) == "A"] + \
              [uea[s,d] for s in ab_staff] + \
              [shuunin_use_a[s,d] for s in shuunin_list]
        model.Add(sum(a_e) == 1)

        # A遅出
        a_l = [x[s,d,"遅"] for s in staff if unit_map.get(s) == "A"] + \
              [ula[s,d] for s in ab_staff]
        model.Add(sum(a_l) == 1)

        # B早出
        b_e = [x[s,d,"早"] for s in staff if unit_map.get(s) == "B"] + \
              [ueb[s,d] for s in ab_staff] + \
              [shuunin_use_b[s,d] for s in shuunin_list]
        model.Add(sum(b_e) == 1)

        # B遅出
        b_l = [x[s,d,"遅"] for s in staff if unit_map.get(s) == "B"] + \
              [ulb[s,d] for s in ab_staff]
        model.Add(sum(b_l) == 1)

        # 夜勤（主任は夜勤なし）
        model.Add(sum(x[s,d,"夜"] for s in staff) == 1)

    # ── 制約6: 夜勤回数 ──
    for s in staff:
        nt = sum(x[s,d,"夜"] for d in range(N))
        model.Add(nt >= nmin_map[s])
        model.Add(nt <= nmax_map[s])
    for s in shuunin_list:
        # 主任は夜勤0
        for d in range(N):
            model.Add(xs[s,d,"夜"] == 0)

    # ── 制約7: 夜勤→翌日（通常職員）──
    # 連続夜勤可 の職員は「夜or×」どちらかを許可
    # 連続夜勤不可 の職員は必ず×
    cn_vars = {}  # cn_vars[s,d]: d日目とd+1日目の連続夜勤フラグ
    for s in staff:
        can_consec = (consec_night_map.get(s, "×") == "○")
        for d in range(N - 1):
            if can_consec:
                # 翌日は×か夜のどちらか（早遅日有は禁止）
                for sh in ["早","遅","日","有"]:
                    model.Add(x[s,d+1,sh] == 0).OnlyEnforceIf(x[s,d,"夜"])
                # 連続夜勤フラグ
                cn = model.NewBoolVar(f"cn_{s}_{d}")
                cn_vars[s,d] = cn
                model.AddBoolAnd([x[s,d,"夜"], x[s,d+1,"夜"]]).OnlyEnforceIf(cn)
                model.AddBoolOr([x[s,d,"夜"].Not(), x[s,d+1,"夜"].Not()]).OnlyEnforceIf(cn.Not())
                # 連続夜勤後は2日×
                if d + 3 < N:
                    model.Add(x[s,d+2,"×"] == 1).OnlyEnforceIf(cn)
                    model.Add(x[s,d+3,"×"] == 1).OnlyEnforceIf(cn)
                elif d + 2 < N:
                    model.Add(x[s,d+2,"×"] == 1).OnlyEnforceIf(cn)
                # 3連続夜勤禁止
                if d + 2 < N:
                    model.Add(x[s,d,"夜"] + x[s,d+1,"夜"] + x[s,d+2,"夜"] <= 2)
            else:
                # 通常: 夜勤→翌日必ず×
                model.Add(x[s,d+1,"×"] == 1).OnlyEnforceIf(x[s,d,"夜"])

    # 主任も夜勤なしなので夜→×は不要だが念のため
    for s in shuunin_list:
        for d in range(N - 1):
            model.Add(xs[s,d+1,"×"] == 1).OnlyEnforceIf(xs[s,d,"夜"])

    # ── 制約8: 遅→翌早禁止 ──
    for s in staff:
        for d in range(N - 1):
            model.Add(x[s,d,"遅"] + x[s,d+1,"早"] <= 1)
    for s in shuunin_list:
        for d in range(N - 1):
            model.Add(xs[s,d,"遅"] + xs[s,d+1,"早"] <= 1)

    # ── 制約9: 希望休前日に夜勤を入れない ──
    for s in staff:
        for date_obj, (sh_type, req_type) in requests.get(s, {}).items():
            if req_type == "希望" and sh_type in ["×","有"]:
                for d, dn in enumerate(days_norm):
                    if dn == date_obj and d > 0:
                        model.Add(x[s,d-1,"夜"] == 0)
                        break

    # ── 制約10: 連勤制限 ──
    for s in staff:
        max_c  = 5 if cont_map[s] == "40h" else 4
        prev_c = count_trailing_consec(prev_month.get(s, []))
        remain = max(0, max_c - prev_c)
        if prev_c > 0 and remain < max_c:
            for w in range(1, min(remain + 2, N + 1)):
                if w > remain:
                    model.Add(sum(x[s,d2,sh2] for d2 in range(w)
                                  for sh2 in ["早","遅","夜","有","日"]) <= remain)
                    break
        for st in range(N - max_c):
            model.Add(sum(x[s,d2,sh2] for d2 in range(st, st+max_c+1)
                          for sh2 in ["早","遅","夜","有","日"]) <= max_c)

    # ── 制約11: 公休数の下限 ──
    for s in staff:
        min_hol = holiday_limits.get(cont_map[s], 8)
        if min_hol > 0:
            model.Add(sum(x[s,d,"×"] for d in range(N)) >= min_hol)

    # ── 制約12: 備考による勤務制限 ──
    for s in all_staff_names:
        allowed = allowed_shifts_map.get(s)
        if allowed is None:
            continue
        forbidden = set(WORK_SHIFTS) - allowed
        var_d = xs if s in shuunin_list else x
        for d in range(N):
            for sh in forbidden:
                req = requests.get(s, {}).get(days_norm[d])
                if req and req[0] == sh and req[1] == "指定":
                    continue
                model.Add(var_d[s,d,sh] == 0)

    # ── 制約13: パート職員に有給を自動割り当てしない ──
    for s in part_staff:
        for d in range(N):
            req = requests.get(s, {}).get(days_norm[d])
            if req and req[0] == "有" and req[1] == "指定":
                pass
            else:
                model.Add(x[s,d,"有"] == 0)

    # ── 制約14: パート職員の週単位勤務日数 ──
    for s in staff:
        if s not in weekly_work_days:
            continue
        target = weekly_work_days[s]
        for week_key in sorted_week_keys:
            didx = week_groups[week_key]
            wv = [x[s,d,sh] for d in didx for sh in ["早","遅","夜","有","日"]]
            if len(didx) == 7:
                model.Add(sum(wv) >= max(0, target - 1))
                model.Add(sum(wv) <= target)
            else:
                model.Add(sum(wv) <= round(target * len(didx) / 7 + 0.5))

    # ── 制約15: 主任は早出か×のみ（有給・遅・夜・日すべて禁止） ──
    for s in shuunin_list:
        for d in range(N):
            for sh in ["遅","夜","日","有"]:
                req = requests.get(s, {}).get(days_norm[d])
                # Shift_Requestsで明示的に指定されている場合のみ例外
                if req and req[0] == sh and req[1] == "指定":
                    continue
                model.Add(xs[s,d,sh] == 0)

    # ======================================================
    # ソフト制約 & 目的関数
    # ======================================================
    penalty_terms = []

    # ── ソフト1: 主任使用日数（最優先で避ける）──
    for s in shuunin_list:
        for d in range(N):
            # 主任が働く日（×以外）にペナルティ
            work_var = model.NewBoolVar(f"sh_work_{s}_{d}")
            model.Add(xs[s,d,"×"] == 0).OnlyEnforceIf(work_var)
            model.Add(xs[s,d,"×"] == 1).OnlyEnforceIf(work_var.Not())
            # 主任早出 = use_a or use_b のどちらか
            penalty_terms.append((xs[s,d,"早"], 200))

    # ── ソフト2: 連続夜勤使用ペナルティ ──
    for (s, d), cn in cn_vars.items():
        penalty_terms.append((cn, 30))

    # ── ソフト3: 公休日数を目標値に近づける（リーダー以外）──
    for s in staff:
        if role_map.get(s, "") == "リーダー":
            continue
        target_off = holiday_limits.get(cont_map[s], 8)
        if target_off <= 0:
            continue
        off_count = model.NewIntVar(0, N, f"off_{s}")
        model.Add(off_count == sum(x[s,d,"×"] for d in range(N)))
        # オーバー分（公休が多すぎる→勤務を増やす）
        over_v  = model.NewIntVar(0, N, f"over_{s}")
        under_v = model.NewIntVar(0, N, f"under_{s}")
        model.Add(over_v  >= off_count - target_off)
        model.Add(over_v  >= 0)
        model.Add(under_v >= target_off - off_count)
        model.Add(under_v >= 0)
        # 目的: オーバーも減らしたいが、アンダー（公休少なすぎ）は許容
        # 公休過多（over）にだけペナルティ（= もっと勤務を入れる）
        penalty_terms.append((over_v,  8))  # 公休が多すぎたら減らす
        penalty_terms.append((under_v, 4))  # 公休が少なすぎても軽ペナルティ

    # ── ソフト4: 早遅の平準化（リーダー以外）──
    non_leader = [s for s in staff if role_map.get(s) != "リーダー"]
    if len(non_leader) >= 2:
        e_vars = []; l_vars = []
        for s in non_leader:
            ev = model.NewIntVar(0, N, f"e_{s}")
            lv = model.NewIntVar(0, N, f"l_{s}")
            model.Add(ev == sum(x[s,d,"早"] for d in range(N)))
            model.Add(lv == sum(x[s,d,"遅"] for d in range(N)))
            e_vars.append(ev); l_vars.append(lv)
        max_e = model.NewIntVar(0, N, "max_e"); min_e = model.NewIntVar(0, N, "min_e")
        max_l = model.NewIntVar(0, N, "max_l"); min_l = model.NewIntVar(0, N, "min_l")
        model.AddMaxEquality(max_e, e_vars); model.AddMinEquality(min_e, e_vars)
        model.AddMaxEquality(max_l, l_vars); model.AddMinEquality(min_l, l_vars)
        diff_e = model.NewIntVar(0, N, "diff_e"); model.Add(diff_e == max_e - min_e)
        diff_l = model.NewIntVar(0, N, "diff_l"); model.Add(diff_l == max_l - min_l)
        penalty_terms.append((diff_e, 5))
        penalty_terms.append((diff_l, 5))

    # ── ソフト5: 勤務間隔（4連続勤務にペナルティ）──
    for s in staff:
        if s in part_with_fixed:
            continue
        for d in range(N - 3):
            work_d = [model.NewBoolVar(f"wd4_{s}_{d}_{k}") for k in range(4)]
            for k in range(4):
                model.Add(sum(x[s,d+k,sh] for sh in ["早","遅","夜","日","有"]) == 1
                          ).OnlyEnforceIf(work_d[k])
                model.Add(sum(x[s,d+k,sh] for sh in ["早","遅","夜","日","有"]) == 0
                          ).OnlyEnforceIf(work_d[k].Not())
            w4_real = model.NewBoolVar(f"w4r_{s}_{d}")
            model.AddBoolAnd(work_d).OnlyEnforceIf(w4_real)
            model.AddBoolOr([w.Not() for w in work_d]).OnlyEnforceIf(w4_real.Not())
            penalty_terms.append((w4_real, 2))

    # ── ソフト6: 同一勤務3連続にペナルティ ──
    for s in staff:
        if s in part_with_fixed:
            continue
        for sh in ["早", "遅"]:
            for d in range(N - 2):
                sc3 = model.NewBoolVar(f"sc3_{s}_{sh}_{d}")
                model.AddBoolAnd([x[s,d,sh], x[s,d+1,sh], x[s,d+2,sh]]).OnlyEnforceIf(sc3)
                model.AddBoolOr([x[s,d,sh].Not(), x[s,d+1,sh].Not(),
                                 x[s,d+2,sh].Not()]).OnlyEnforceIf(sc3.Not())
                penalty_terms.append((sc3, 3))

    # ── 目的関数 ──
    obj_terms = []
    for var, coef in penalty_terms:
        obj_terms.append(var * coef)
    if obj_terms:
        model.Minimize(sum(obj_terms))

    # ======================================================
    # ソルバー実行
    # ======================================================
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 300
    solver.parameters.num_search_workers  = 8
    status = solver.Solve(model)

    if status not in (cp_model.FEASIBLE, cp_model.OPTIMAL):
        raise Exception(
            "条件を満たすシフト表が見つかりませんでした。\n"
            "希望シフト・夜勤回数・公休数の設定を見直してください。"
        )

    # ── 結果組み立て ──
    result = {}
    for s in staff:
        result[s] = {}
        for d in range(N):
            for sh in ALL_SHIFTS:
                if solver.Value(x[s,d,sh]) == 1:
                    result[s][d] = sh
                    break

    for s in shuunin_list:
        result[s] = {}
        for d in range(N):
            for sh in ALL_SHIFTS:
                if solver.Value(xs[s,d,sh]) == 1:
                    result[s][d] = sh
                    break

    # A・B職員ユニット割り当て
    ab_unit_result = {}
    for s in ab_staff:
        ab_unit_result[s] = {}
        for d in range(N):
            sh = result[s][d]
            if sh == "早":
                ab_unit_result[s][d] = "A" if solver.Value(uea[s,d]) == 1 else "B"
            elif sh == "遅":
                ab_unit_result[s][d] = "A" if solver.Value(ula[s,d]) == 1 else "B"
            else:
                ab_unit_result[s][d] = None

    # 主任がどのユニットに入ったか
    shuunin_unit_result = {}
    for s in shuunin_list:
        shuunin_unit_result[s] = {}
        for d in range(N):
            ua = solver.Value(shuunin_use_a[s,d])
            ub = solver.Value(shuunin_use_b[s,d])
            if ua:
                shuunin_unit_result[s][d] = "A"
            elif ub:
                shuunin_unit_result[s][d] = "B"
            else:
                shuunin_unit_result[s][d] = None

    return (result, staff, shuunin_list, unit_map, cont_map, role_map,
            days_norm, requests, ab_unit_result, shuunin_unit_result)


# ========================================================
# Excel 書き出し
# ========================================================
def write_shift_result(result, staff, shuunin_list, unit_map, cont_map, role_map,
                       days_norm, requests, ab_unit_result, shuunin_unit_result,
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
    SUMMARY_HDRS   = ["早出","遅出","日勤","夜勤","公休"]

    all_disp_staff = shuunin_list + staff   # 主任を先頭に
    STAFF_START_ROW  = 4
    SHUUNIN_SEP_ROW  = STAFF_START_ROW + len(shuunin_list)
    SUMMARY_ROW_BASE = STAFF_START_ROW + len(all_disp_staff) + 1

    # ── ヘッダー ──
    ws.cell(1, 1, "作成月")
    ws.cell(1, 2, days_norm[0].strftime("%Y年%m月"))
    ws.cell(2, 2, "曜日")
    ws.cell(3, 1, "ユニット")
    ws.cell(3, 2, "職員名")

    for i, d in enumerate(days_norm):
        col = DATE_START_COL + i
        ws.cell(1, col, d.day).alignment = Alignment(horizontal="center")
        wd_cell = ws.cell(2, col, weekday_ja[d.weekday()])
        wd_cell.alignment = Alignment(horizontal="center")
        if d.weekday() == 5:
            wd_cell.fill = PatternFill("solid", fgColor="CCE5FF")
        elif d.weekday() == 6:
            wd_cell.fill = PatternFill("solid", fgColor="FFCCCC")

    for k, h in enumerate(SUMMARY_HDRS):
        c = ws.cell(3, SUMMARY_COL + k, h)
        c.fill = YELLOW_FILL
        c.alignment = Alignment(horizontal="center")
    ws.cell(3, 1).fill = YELLOW_FILL
    ws.cell(3, 2).fill = YELLOW_FILL

    # ── 主任行（上部に表示）──
    for idx, s in enumerate(shuunin_list):
        row = STAFF_START_ROW + idx
        u_label = "主任"
        ws.cell(row, 1, u_label).alignment = Alignment(horizontal="center")
        ws.cell(row, 2, s).alignment = Alignment(horizontal="center")
        ws.cell(row, 1).fill = BLUE_FILL
        ws.cell(row, 2).fill = BLUE_FILL

        for d in range(N):
            col  = DATE_START_COL + d
            sh   = result[s][d]
            cell = ws.cell(row, col, sh)
            cell.alignment = Alignment(horizontal="center")
            date_obj = days_norm[d]
            # 主任が使われた日は青色
            su_r = shuunin_unit_result.get(s, {}).get(d)
            if sh == "早" and su_r:
                cell.fill = BLUE_FILL
            elif s in requests and date_obj in requests[s]:
                _, rtype = requests[s][date_obj]
                if rtype == "希望":
                    cell.fill = PINK_FILL
                elif rtype == "指定":
                    cell.fill = GREEN_FILL

        ds  = get_column_letter(DATE_START_COL)
        de  = get_column_letter(DATE_START_COL + N - 1)
        rng = f"{ds}{row}:{de}{row}"
        ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"早")')
        ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"遅")')
        ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"日")')
        ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')
        ws.cell(row, SUMMARY_COL + 4, f'=COUNTIF({rng},"×")')

    # 主任と一般職員の区切り線
    if shuunin_list:
        sep_row = SHUUNIN_SEP_ROW
        for col in range(1, SUMMARY_COL + len(SUMMARY_HDRS)):
            ws.cell(sep_row, col).fill = PatternFill("solid", fgColor="E0E0E0")

    # ── 一般職員行 ──
    def unit_order(s):
        u = unit_map.get(s, "")
        if u == "A":    return 0
        if u == "A・B": return 1
        return 2
    sorted_staff = sorted(staff, key=unit_order)

    for idx, s in enumerate(sorted_staff):
        row = SHUUNIN_SEP_ROW + idx + (1 if shuunin_list else 0)
        ws.cell(row, 1, unit_map.get(s, "")).alignment = Alignment(horizontal="center")
        ws.cell(row, 2, s).alignment = Alignment(horizontal="center")

        for d in range(N):
            col  = DATE_START_COL + d
            sh   = result[s][d]
            cell = ws.cell(row, col, sh)
            cell.alignment = Alignment(horizontal="center")
            date_obj = days_norm[d]
            if s in requests and date_obj in requests[s]:
                _, rtype = requests[s][date_obj]
                if rtype == "希望":
                    cell.fill = PINK_FILL
                elif rtype == "指定":
                    cell.fill = GREEN_FILL

        ds  = get_column_letter(DATE_START_COL)
        de  = get_column_letter(DATE_START_COL + N - 1)
        rng = f"{ds}{row}:{de}{row}"
        ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"早")')
        ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"遅")')
        ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"日")')
        ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')
        ws.cell(row, SUMMARY_COL + 4, f'=COUNTIF({rng},"×")')

    # ── 日別集計行 ──
    ab_staff_local = [s for s in staff if unit_map.get(s) == "A・B"]
    label_names = ["A早出","B早出","A遅出","B遅出","夜勤"]
    for k, lbl in enumerate(label_names):
        r = SUMMARY_ROW_BASE + k
        c = ws.cell(r, 2, lbl)
        c.fill = GRAY_FILL
        c.alignment = Alignment(horizontal="center")

    for i in range(N):
        col = DATE_START_COL + i
        cnt_ae = (sum(1 for s in staff if unit_map.get(s)=="A" and result[s][i]=="早") +
                  sum(1 for s in ab_staff_local if ab_unit_result.get(s,{}).get(i)=="A" and result[s][i]=="早") +
                  sum(1 for s in shuunin_list if shuunin_unit_result.get(s,{}).get(i)=="A" and result[s][i]=="早"))
        cnt_be = (sum(1 for s in staff if unit_map.get(s)=="B" and result[s][i]=="早") +
                  sum(1 for s in ab_staff_local if ab_unit_result.get(s,{}).get(i)=="B" and result[s][i]=="早") +
                  sum(1 for s in shuunin_list if shuunin_unit_result.get(s,{}).get(i)=="B" and result[s][i]=="早"))
        cnt_al = (sum(1 for s in staff if unit_map.get(s)=="A" and result[s][i]=="遅") +
                  sum(1 for s in ab_staff_local if ab_unit_result.get(s,{}).get(i)=="A" and result[s][i]=="遅"))
        cnt_bl = (sum(1 for s in staff if unit_map.get(s)=="B" and result[s][i]=="遅") +
                  sum(1 for s in ab_staff_local if ab_unit_result.get(s,{}).get(i)=="B" and result[s][i]=="遅"))
        cnt_nt = sum(1 for s in staff if result[s][i]=="夜")
        for k, v in enumerate([cnt_ae, cnt_be, cnt_al, cnt_bl, cnt_nt]):
            ws.cell(SUMMARY_ROW_BASE + k, col, v).alignment = Alignment(horizontal="center")

    # 列幅
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 8
    for i in range(N):
        ws.column_dimensions[get_column_letter(DATE_START_COL + i)].width = 4
    for k in range(len(SUMMARY_HDRS)):
        ws.column_dimensions[get_column_letter(SUMMARY_COL + k)].width = 6

    wb.save(output_path)


# ========================================================
# Web UI
# ========================================================
HTML_CONTENT = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>シフト表自動作成アプリ v4.0</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;display:flex;justify-content:center;align-items:flex-start;padding:30px 20px}
.card{background:#fff;padding:40px;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-width:960px;width:100%}
h1{color:#667eea;font-size:1.9em;text-align:center;margin-bottom:6px}
.ver{text-align:center;color:#764ba2;font-weight:bold;margin-bottom:4px;font-size:.9em}
.sub{text-align:center;color:#888;margin-bottom:20px;font-size:.85em}
.sec-title{font-weight:bold;color:#333;margin-bottom:10px;font-size:1em;border-left:4px solid #667eea;padding-left:10px;margin-top:18px}
.rules{background:#f8f9fa;padding:14px 20px;border-radius:10px;margin-bottom:14px}
.rules ul{list-style:none}
.rules li{padding:4px 0;border-bottom:1px solid #eee;font-size:.86em;color:#555}
.rules li:last-child{border-bottom:none}
.badge{display:inline-block;background:#667eea;color:#fff;padding:1px 7px;border-radius:10px;font-size:.75em;margin-left:4px;vertical-align:middle}
.badge.new{background:#e74c3c}
.note{background:#fff8e1;border-left:4px solid #ffc107;padding:12px 16px;border-radius:5px;margin-bottom:18px;font-size:.86em;color:#555;line-height:1.7}
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
.pmsg{color:#667eea;font-size:.9em;line-height:1.6}
.ok{display:none;background:#d4edda;border:1px solid #c3e6cb;color:#155724;padding:16px;border-radius:10px;margin-top:16px;text-align:center}
.dl-btn{display:inline-block;background:#28a745;color:#fff;padding:10px 28px;text-decoration:none;border-radius:20px;margin-top:10px;font-size:.95em}
.dl-btn:hover{background:#218838}
.err{display:none;background:#f8d7da;border:1px solid #f5c6cb;color:#721c24;padding:14px;border-radius:10px;margin-top:16px;word-break:break-all;white-space:pre-wrap;font-size:.88em}
.legend{display:flex;gap:14px;margin-top:14px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px;font-size:.82em;color:#555}
.sw{width:16px;height:16px;border-radius:3px;border:1px solid #ccc}
.c-pink{background:#FFB6C1}.c-green{background:#90EE90}.c-blue{background:#BDD7EE}
</style>
</head>
<body>
<div class="card">
  <h1>📅 シフト表自動作成アプリ</h1>
  <p class="ver">Version 4.0</p>
  <p class="sub">Excelをアップロードするだけで最適なシフト表を自動生成</p>

  <div class="sec-title">🔒 適用される制約・ルール</div>
  <div class="rules"><ul>
    <li>✅ ユニットA/B：毎日<strong>早出1・遅出1</strong>（A・B兼務職員はどちらか一方にカウント）</li>
    <li>✅ 夜勤：毎日1名（個人の最少〜最高回数を厳守）</li>
    <li>✅ 40h→最大5連勤 / 32h・パート→最大4連勤（前月継続分を考慮）</li>
    <li>✅ 夜勤→翌日×、遅出→翌日早出禁止</li>
    <li>✅ 希望休の前日夜勤禁止、パート職員に有給を自動割り当てしない</li>
    <li>✅ Staff_Masterの備考（早出のみ・週N日勤務・夜勤なし等）を厳守</li>
    <li>✅ 固定公休（曜日指定）対応</li>
    <li>✅ <strong>公休日数をなるべく指定日数に近づける</strong>（リーダー以外）<span class="badge new">NEW</span></li>
    <li>✅ <strong>連続夜勤</strong>：Staff_Masterで○の職員のみ緊急時に「夜夜××」を許可<span class="badge new">NEW</span></li>
    <li>✅ <strong>勤務間隔</strong>：なるべく3〜4日に1回は休みになるよう配慮<span class="badge new">NEW</span></li>
    <li>✅ <strong>同一勤務の連続を回避</strong>：「早早早」「遅遅遅」をなるべく避ける<span class="badge new">NEW</span></li>
    <li>✅ <strong>主任</strong>：本来の職員では組めないときのみ早出で補完（通常は使わない）<span class="badge new">NEW</span></li>
  </ul></div>

  <div class="note">
    <strong>📋 必要なシート：</strong> Staff_Master / Settings / Shift_Requests / Prev_Month / shift_result<br>
    <strong>【連続夜勤】</strong> Staff_Masterの「連続夜勤」欄に「○」を記入した職員のみ、どうしても夜勤が組めない場合に「夜夜××」が発生します。<br>
    <strong>【主任】</strong> ユニット欄が空欄の職員は主任扱いになります。緊急時のみ早出でAまたはBユニットを補完します（Excel上で青色表示）。<br>
    <strong>【公休日数】</strong> リーダー以外の公休数は、Settingsで指定した日数に近づくよう自動調整します。
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
    <p class="pmsg" id="pmsg">生成中… <strong>0秒</strong> 経過<br>最大5分かかる場合があります。そのままお待ちください。</p>
  </div>
  <div class="ok" id="ok">
    <p>✅ シフト表の生成が完了しました！</p>
    <a href="#" id="dl" class="dl-btn">📥 Shift_Result.xlsx をダウンロード</a>
  </div>
  <div class="err" id="er"></div>

  <div class="legend">
    <div class="legend-item"><div class="sw c-pink"></div>希望休・有給（希望）</div>
    <div class="legend-item"><div class="sw c-green"></div>勤務指定（指定）</div>
    <div class="legend-item"><div class="sw c-blue"></div>主任補完（緊急使用）</div>
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
  const f=e.dataTransfer.files;
  if(f[0]){const dt=new DataTransfer();dt.items.add(f[0]);fi.files=dt.files;fname.textContent='📄 '+f[0].name;}
});
drop.addEventListener('click',()=>fi.click());
let elapsed=0,timer=null;
function startTimer(){elapsed=0;timer=setInterval(()=>{elapsed++;pmsg.innerHTML='生成中… <strong>'+elapsed+'秒</strong> 経過<br>最大5分かかる場合があります。そのままお待ちください。';},1000);}
function stopTimer(){if(timer){clearInterval(timer);timer=null;}}
frm.onsubmit=async e=>{
  e.preventDefault();
  if(!fi.files[0]){alert('ファイルを選択してください');return;}
  const fd=new FormData();fd.append('file',fi.files[0]);
  sw.style.display='block';ok.style.display='none';er.style.display='none';go.disabled=true;
  startTimer();
  try{
    const res=await fetch('/generate-shift',{method:'POST',body:fd});
    stopTimer();
    if(res.ok){
      const blob=await res.blob();
      dl.href=URL.createObjectURL(blob);dl.download='Shift_Result.xlsx';
      sw.style.display='none';ok.style.display='block';
    }else{
      const j=await res.json().catch(()=>({}));
      throw new Error(j.detail||'サーバーエラーが発生しました');
    }
  }catch(ex){
    stopTimer();sw.style.display='none';er.style.display='block';
    er.textContent='❌ エラー:\\n'+ex.message;
  }finally{go.disabled=false;}
};
</script>
</body>
</html>"""


# ========================================================
# FastAPI Routes
# ========================================================
@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(content=HTML_CONTENT)

@app.get("/health")
async def health():
    return {"status": "ok", "version": "4.0"}

@app.post("/generate-shift")
async def generate(file: UploadFile = File(...)):
    uid     = str(uuid.uuid4())
    in_p    = os.path.join(TEMP_DIR, f"in_{uid}.xlsx")
    out_p   = os.path.join(TEMP_DIR, f"out_{uid}.xlsx")
    try:
        with open(in_p, "wb") as f:
            shutil.copyfileobj(file.file, f)
        (result, staff, shuunin_list, unit_map, cont_map, role_map,
         days_norm, requests, ab_unit_result, shuunin_unit_result) = generate_shift(in_p)
        write_shift_result(
            result, staff, shuunin_list, unit_map, cont_map, role_map,
            days_norm, requests, ab_unit_result, shuunin_unit_result,
            in_p, out_p)
        return FileResponse(
            out_p, filename="Shift_Result.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        try: os.remove(in_p)
        except: pass


# ========================================================
# スタンドアロン起動
# ========================================================
if __name__ == "__main__":
    import uvicorn, webbrowser, threading, time

    def open_browser():
        time.sleep(2.0)
        webbrowser.open("http://localhost:8000")

    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    if os.environ.get("AUTO_BROWSER", "1") == "1" and port == 8000:
        threading.Thread(target=open_browser, daemon=True).start()

    print("=" * 50)
    print(" シフト表自動作成アプリ v4.0")
    print(f" http://localhost:{port}")
    print("=" * 50)
    uvicorn.run("main:app", host=host, port=port, reload=False)
