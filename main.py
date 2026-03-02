"""
シフト表自動作成アプリ v5.0
新機能:
  - Staff_Masterのユニット列をA/Bのみに制限
  - ユニット兼務（◯✕）列を追加：兼務職員が他ユニット勤務時はA早/B早/A遅/B遅で表示
  - 集計行にCOUNTIF数式（A早/B早/A遅/B遅/夜勤）を使用
  - 主任のシフトをA早/B早で表示
  - favicon.pngをファビコンとして使用
  - .xlsmファイルの選択・アップロードに対応
  - WebUIを全面リデザイン（ロゴ・アニメーション強化）
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
import pandas as pd
import shutil, os, uuid, re, base64, pathlib
from ortools.sat.python import cp_model
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

app = FastAPI(title="シフト表自動作成アプリ v5.0")
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
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FILL2= PatternFill("solid", fgColor="5B9BD5")
A_UNIT_FILL = PatternFill("solid", fgColor="DEEAF1")   # Aユニット薄青
B_UNIT_FILL = PatternFill("solid", fgColor="E2EFDA")   # Bユニット薄緑

WEEKDAY_MAP = {
    "月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6,
    "月曜": 0, "火曜": 1, "水曜": 2, "木曜": 3, "金曜": 4, "土曜": 5, "日曜": 6,
}

# ── favicon 読み込み ──
FAVICON_B64 = ""
_favicon_path = pathlib.Path(__file__).parent / "favicon.png"
if _favicon_path.exists():
    with open(_favicon_path, "rb") as _f:
        FAVICON_B64 = base64.b64encode(_f.read()).decode()


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

    # ── ユニット兼務列の読み込み ──
    kanmu_col = next((c for c in staff_df.columns if "兼務" in str(c)), None)
    if kanmu_col:
        kanmu_map = dict(zip(staff_df["職員名"], staff_df[kanmu_col].astype(str).str.strip()))
    else:
        # 後方互換: ユニット列が "A・B" の場合も兼務とみなす
        kanmu_map = {}
        for s in all_staff_names:
            u = str(unit_map.get(s, "")).strip()
            if u == "A・B":
                kanmu_map[s] = "○"
            else:
                kanmu_map[s] = "×"

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

    # 主任の識別（ユニット欄がnull/nanの場合）
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

    # ── 兼務職員（ユニット兼務=○）──
    ab_staff = [s for s in staff if kanmu_map.get(s, "×") == "○"]
    ab_staff_set = set(ab_staff)

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
    xs = {}
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
            model.Add(shuunin_use_a[s,d] + shuunin_use_b[s,d] <= xs[s,d,"早"])
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
    # 兼務職員はuea/ueb/ula/ulbで管理（固定A/Bリストから除外）
    for d in range(N):
        # A早出（固定Aスタッフ + 兼務→Aユニット + 主任→Aユニット）
        a_e = ([x[s,d,"早"] for s in staff if unit_map.get(s) == "A" and s not in ab_staff_set] +
               [uea[s,d] for s in ab_staff] +
               [shuunin_use_a[s,d] for s in shuunin_list])
        model.Add(sum(a_e) == 1)

        # A遅出（固定Aスタッフ + 兼務→Aユニット）
        a_l = ([x[s,d,"遅"] for s in staff if unit_map.get(s) == "A" and s not in ab_staff_set] +
               [ula[s,d] for s in ab_staff])
        model.Add(sum(a_l) == 1)

        # B早出（固定Bスタッフ + 兼務→Bユニット + 主任→Bユニット）
        b_e = ([x[s,d,"早"] for s in staff if unit_map.get(s) == "B" and s not in ab_staff_set] +
               [ueb[s,d] for s in ab_staff] +
               [shuunin_use_b[s,d] for s in shuunin_list])
        model.Add(sum(b_e) == 1)

        # B遅出（固定Bスタッフ + 兼務→Bユニット）
        b_l = ([x[s,d,"遅"] for s in staff if unit_map.get(s) == "B" and s not in ab_staff_set] +
               [ulb[s,d] for s in ab_staff])
        model.Add(sum(b_l) == 1)

        # 夜勤（主任は夜勤なし）
        model.Add(sum(x[s,d,"夜"] for s in staff) == 1)

    # ── 制約6: 夜勤回数 ──
    for s in staff:
        nt = sum(x[s,d,"夜"] for d in range(N))
        model.Add(nt >= nmin_map[s])
        model.Add(nt <= nmax_map[s])
    for s in shuunin_list:
        for d in range(N):
            model.Add(xs[s,d,"夜"] == 0)

    # ── 制約7: 夜勤→翌日 ──
    cn_vars = {}
    for s in staff:
        can_consec = (consec_night_map.get(s, "×") == "○")
        for d in range(N - 1):
            if can_consec:
                for sh in ["早","遅","日","有"]:
                    model.Add(x[s,d+1,sh] == 0).OnlyEnforceIf(x[s,d,"夜"])
                cn = model.NewBoolVar(f"cn_{s}_{d}")
                cn_vars[s,d] = cn
                model.AddBoolAnd([x[s,d,"夜"], x[s,d+1,"夜"]]).OnlyEnforceIf(cn)
                model.AddBoolOr([x[s,d,"夜"].Not(), x[s,d+1,"夜"].Not()]).OnlyEnforceIf(cn.Not())
                if d + 3 < N:
                    model.Add(x[s,d+2,"×"] == 1).OnlyEnforceIf(cn)
                    model.Add(x[s,d+3,"×"] == 1).OnlyEnforceIf(cn)
                elif d + 2 < N:
                    model.Add(x[s,d+2,"×"] == 1).OnlyEnforceIf(cn)
                if d + 2 < N:
                    model.Add(x[s,d,"夜"] + x[s,d+1,"夜"] + x[s,d+2,"夜"] <= 2)
            else:
                model.Add(x[s,d+1,"×"] == 1).OnlyEnforceIf(x[s,d,"夜"])

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

    # ── 制約15: 主任は早出か×のみ ──
    for s in shuunin_list:
        for d in range(N):
            for sh in ["遅","夜","日","有"]:
                req = requests.get(s, {}).get(days_norm[d])
                if req and req[0] == sh and req[1] == "指定":
                    continue
                model.Add(xs[s,d,sh] == 0)

    # ======================================================
    # ソフト制約 & 目的関数
    # ======================================================
    penalty_terms = []

    # ── ソフト1: 主任使用ペナルティ ──
    for s in shuunin_list:
        for d in range(N):
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
        over_v  = model.NewIntVar(0, N, f"over_{s}")
        under_v = model.NewIntVar(0, N, f"under_{s}")
        model.Add(over_v  >= off_count - target_off)
        model.Add(over_v  >= 0)
        model.Add(under_v >= target_off - off_count)
        model.Add(under_v >= 0)
        penalty_terms.append((over_v,  8))
        penalty_terms.append((under_v, 4))

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

    # 兼務職員ユニット割り当て
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

    # 主任ユニット
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
            days_norm, requests, ab_unit_result, shuunin_unit_result, kanmu_map)


# ========================================================
# Excel 書き出し
# ========================================================
def write_shift_result(result, staff, shuunin_list, unit_map, cont_map, role_map,
                       days_norm, requests, ab_unit_result, shuunin_unit_result,
                       kanmu_map, input_path, output_path):
    """
    ユニット付き表示:
      - 固定A職員の早/遅 → A早/A遅
      - 固定B職員の早/遅 → B早/B遅
      - 兼務職員の早/遅  → 割り当てユニット+早/遅 (例: B早, A遅)
      - 主任の早         → A早 or B早
      - 夜/×/有/日      → そのまま
    """
    wb = load_workbook(input_path, keep_vba=True if input_path.endswith(".xlsm") else False)
    if "shift_result" in wb.sheetnames:
        del wb["shift_result"]
    ws = wb.create_sheet("shift_result")

    N = len(days_norm)
    weekday_ja = ["月","火","水","木","金","土","日"]
    DATE_START_COL = 3
    SUMMARY_COL    = DATE_START_COL + N
    # 個人サマリー列: 早出/遅出/日勤/夜勤/公休
    SUMMARY_HDRS   = ["早出","遅出","日勤","夜勤","公休"]

    all_disp_staff = shuunin_list + staff
    STAFF_START_ROW  = 4
    SHUUNIN_SEP_ROW  = STAFF_START_ROW + len(shuunin_list)
    # sorted_staffを先に計算（COUNTIFレンジ計算のため）
    def unit_order(s):
        u = unit_map.get(s, "")
        k = kanmu_map.get(s, "×")
        if u == "A" and k != "○": return 0
        if k == "○": return 1
        if u == "B": return 2
        return 3
    sorted_staff = sorted(staff, key=unit_order)
    LAST_STAFF_ROW = SHUUNIN_SEP_ROW + len(sorted_staff)
    SUMMARY_ROW_BASE = LAST_STAFF_ROW + 2  # 空白1行を挟む

    # ── ユニット付きシフト文字列を返すヘルパー ──
    def display_val(s, d):
        sh = result[s].get(d, "×")
        if sh not in ("早", "遅"):
            return sh
        if s in shuunin_list:
            unit = shuunin_unit_result.get(s, {}).get(d)
            return (unit + sh) if unit else sh
        elif kanmu_map.get(s, "×") == "○":
            unit = ab_unit_result.get(s, {}).get(d)
            return (unit + sh) if unit else sh
        else:
            unit = unit_map.get(s, "")
            return (unit + sh) if unit in ("A", "B") else sh

    # ── セル色決定 ──
    def cell_fill(s, d):
        date_obj = days_norm[d]
        if s in requests and date_obj in requests[s]:
            _, rtype = requests[s][date_obj]
            if rtype == "希望":
                return PINK_FILL
            elif rtype == "指定":
                return GREEN_FILL
        if s in shuunin_list:
            unit = shuunin_unit_result.get(s, {}).get(d)
            sh   = result[s].get(d, "×")
            if sh == "早" and unit:
                return BLUE_FILL
        return None

    # ── ヘッダー行 ──
    ws.cell(1, 1, "作成月").font = Font(bold=True)
    ws.cell(1, 2, days_norm[0].strftime("%Y年%m月"))
    ws.cell(2, 2, "曜日").alignment = Alignment(horizontal="center")
    ws.cell(3, 1, "ユニット").alignment = Alignment(horizontal="center")
    ws.cell(3, 2, "職員名").alignment  = Alignment(horizontal="center")
    ws.cell(3, 1).fill = HEADER_FILL; ws.cell(3, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(3, 2).fill = HEADER_FILL; ws.cell(3, 2).font = Font(bold=True, color="FFFFFF")

    for i, d in enumerate(days_norm):
        col = DATE_START_COL + i
        c1 = ws.cell(1, col, d.day)
        c1.alignment = Alignment(horizontal="center")
        c1.font = Font(bold=True)
        wd_cell = ws.cell(2, col, weekday_ja[d.weekday()])
        wd_cell.alignment = Alignment(horizontal="center")
        if d.weekday() == 5:
            wd_cell.fill = PatternFill("solid", fgColor="CCE5FF")
        elif d.weekday() == 6:
            wd_cell.fill = PatternFill("solid", fgColor="FFCCCC")
        # 日付ヘッダー行に薄い枠
        ws.cell(3, col).fill = HEADER_FILL2
        ws.cell(3, col).font = Font(color="FFFFFF")

    for k, h in enumerate(SUMMARY_HDRS):
        c = ws.cell(3, SUMMARY_COL + k, h)
        c.fill = YELLOW_FILL
        c.alignment = Alignment(horizontal="center")
        c.font = Font(bold=True)

    # ── 主任行 ──
    for idx, s in enumerate(shuunin_list):
        row = STAFF_START_ROW + idx
        ws.cell(row, 1, "主任").alignment = Alignment(horizontal="center")
        ws.cell(row, 1).fill = BLUE_FILL
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 2, s).alignment = Alignment(horizontal="center")
        ws.cell(row, 2).fill = BLUE_FILL

        for d in range(N):
            col  = DATE_START_COL + d
            val  = display_val(s, d)
            cell = ws.cell(row, col, val)
            cell.alignment = Alignment(horizontal="center")
            f = cell_fill(s, d)
            if f:
                cell.fill = f

        # 個人COUNTIF集計
        ds  = get_column_letter(DATE_START_COL)
        de  = get_column_letter(DATE_START_COL + N - 1)
        rng = f"{ds}{row}:{de}{row}"
        ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"A早")+COUNTIF({rng},"B早")')
        ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"A遅")+COUNTIF({rng},"B遅")')
        ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"日")')
        ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')
        ws.cell(row, SUMMARY_COL + 4, f'=COUNTIF({rng},"×")')
        for k in range(len(SUMMARY_HDRS)):
            ws.cell(row, SUMMARY_COL + k).alignment = Alignment(horizontal="center")

    # 主任と一般職員の区切り線
    if shuunin_list:
        for col in range(1, SUMMARY_COL + len(SUMMARY_HDRS)):
            ws.cell(SHUUNIN_SEP_ROW, col).fill = PatternFill("solid", fgColor="E0E0E0")

    # ── 一般職員行 ──
    for idx, s in enumerate(sorted_staff):
        row = SHUUNIN_SEP_ROW + idx + (1 if shuunin_list else 0)
        u   = unit_map.get(s, "")
        k   = kanmu_map.get(s, "×")
        # ユニット表示
        if k == "○":
            u_label = f"{u}兼"
        else:
            u_label = u
        uc = ws.cell(row, 1, u_label)
        uc.alignment = Alignment(horizontal="center")
        if u == "A":
            uc.fill = A_UNIT_FILL
        elif u == "B":
            uc.fill = B_UNIT_FILL

        nc = ws.cell(row, 2, s)
        nc.alignment = Alignment(horizontal="center")
        if u == "A":
            nc.fill = A_UNIT_FILL
        elif u == "B":
            nc.fill = B_UNIT_FILL

        for d in range(N):
            col  = DATE_START_COL + d
            val  = display_val(s, d)
            cell = ws.cell(row, col, val)
            cell.alignment = Alignment(horizontal="center")
            f = cell_fill(s, d)
            if f:
                cell.fill = f

        # 個人COUNTIF集計
        ds  = get_column_letter(DATE_START_COL)
        de  = get_column_letter(DATE_START_COL + N - 1)
        rng = f"{ds}{row}:{de}{row}"
        # 早出: A早+B早（兼務でない固定スタッフはAor B早のみ）
        ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"A早")+COUNTIF({rng},"B早")')
        ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"A遅")+COUNTIF({rng},"B遅")')
        ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"日")')
        ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')
        ws.cell(row, SUMMARY_COL + 4, f'=COUNTIF({rng},"×")')
        for k2 in range(len(SUMMARY_HDRS)):
            ws.cell(row, SUMMARY_COL + k2).alignment = Alignment(horizontal="center")

    # ── 日別集計行（COUNTIF数式）──
    # COUNTIFレンジ: 全スタッフ行（STAFF_START_ROW ～ LAST_STAFF_ROW）
    daily_labels = ["A早出","B早出","A遅出","B遅出","夜勤"]
    daily_values = ["A早",  "B早",  "A遅",  "B遅",  "夜" ]
    daily_fills  = [A_UNIT_FILL, B_UNIT_FILL, A_UNIT_FILL, B_UNIT_FILL, GRAY_FILL]

    for k, (lbl, fill) in enumerate(zip(daily_labels, daily_fills)):
        r = SUMMARY_ROW_BASE + k
        c = ws.cell(r, 2, lbl)
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
        c.font = Font(bold=True)

    for i in range(N):
        col = DATE_START_COL + i
        col_letter = get_column_letter(col)
        # スタッフ全行のレンジ（区切り行を含むが空白なのでCOUNTIFに影響なし）
        cnt_range = f"{col_letter}{STAFF_START_ROW}:{col_letter}{LAST_STAFF_ROW}"
        for k, (_, dv) in enumerate(zip(daily_labels, daily_values)):
            r = SUMMARY_ROW_BASE + k
            c = ws.cell(r, col, f'=COUNTIF({cnt_range},"{dv}")')
            c.alignment = Alignment(horizontal="center")

    # ── 列幅 ──
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 10
    for i in range(N):
        ws.column_dimensions[get_column_letter(DATE_START_COL + i)].width = 5
    for k in range(len(SUMMARY_HDRS)):
        ws.column_dimensions[get_column_letter(SUMMARY_COL + k)].width = 8

    wb.save(output_path)


# ========================================================
# Web UI HTML
# ========================================================
_favicon_tag = (f'<link rel="icon" type="image/png" href="data:image/png;base64,{FAVICON_B64}">'
                if FAVICON_B64 else '<link rel="icon" href="/favicon.png">')

HTML_CONTENT = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>シフト表自動作成アプリ v5.0</title>
{_favicon_tag}
<style>
/* ── リセット・ベース ── */
*{{margin:0;padding:0;box-sizing:border-box}}
:root{{
  --primary:#2563EB;--primary-dark:#1D4ED8;--primary-light:#DBEAFE;
  --accent:#7C3AED;--accent-light:#EDE9FE;
  --success:#059669;--warn:#D97706;--danger:#DC2626;
  --gray-50:#F9FAFB;--gray-100:#F3F4F6;--gray-200:#E5E7EB;--gray-600:#4B5563;--gray-800:#1F2937;
  --radius:16px;--shadow:0 25px 50px rgba(0,0,0,.18);
}}
body{{
  font-family:'Segoe UI','Noto Sans JP',sans-serif;
  background:linear-gradient(135deg,#0f2027,#203a43,#2c5364);
  min-height:100vh;display:flex;flex-direction:column;align-items:center;
  padding:24px 16px 48px;
}}

/* ── ヘッダー ── */
.app-header{{
  width:100%;max-width:900px;display:flex;align-items:center;
  justify-content:space-between;margin-bottom:28px;
  animation:slideDown .6s ease both;
}}
.logo-block{{display:flex;align-items:center;gap:14px}}
.logo-icon{{
  width:60px;height:60px;background:linear-gradient(135deg,var(--primary),var(--accent));
  border-radius:16px;display:flex;align-items:center;justify-content:center;
  box-shadow:0 8px 24px rgba(37,99,235,.45);
  transition:transform .3s;
}}
.logo-icon:hover{{transform:rotate(8deg) scale(1.08)}}
.logo-icon svg{{width:34px;height:34px;fill:#fff}}
.logo-text h1{{font-size:1.55em;font-weight:800;color:#fff;letter-spacing:-.5px;line-height:1.1}}
.logo-text p{{font-size:.78em;color:rgba(255,255,255,.7);margin-top:2px}}
.ver-badge{{
  background:linear-gradient(135deg,var(--primary),var(--accent));
  color:#fff;padding:6px 14px;border-radius:20px;font-size:.78em;font-weight:700;
  box-shadow:0 4px 12px rgba(124,58,237,.4);white-space:nowrap;
}}

/* ── メインカード ── */
.card{{
  background:rgba(255,255,255,.97);
  border-radius:var(--radius);box-shadow:var(--shadow);
  max-width:900px;width:100%;overflow:hidden;
  animation:fadeUp .7s .15s ease both;
}}

/* ── セクション ── */
.section{{padding:28px 32px;border-bottom:1px solid var(--gray-200)}}
.section:last-child{{border-bottom:none}}
.sec-title{{
  display:flex;align-items:center;gap:10px;
  font-size:1.02em;font-weight:700;color:var(--gray-800);
  margin-bottom:14px;
}}
.sec-title .icon{{
  width:32px;height:32px;background:var(--primary-light);
  border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:1em;
}}

/* ── 制約リスト ── */
.rules-grid{{
  display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));
  gap:8px;
}}
.rule-item{{
  display:flex;align-items:flex-start;gap:8px;
  background:var(--gray-50);border-radius:8px;
  padding:8px 12px;font-size:.85em;color:var(--gray-600);
  border-left:3px solid var(--gray-200);
  transition:border-color .2s,background .2s;
}}
.rule-item:hover{{background:var(--primary-light);border-left-color:var(--primary)}}
.rule-item .chk{{color:var(--success);font-weight:700;flex-shrink:0}}
.badge{{
  display:inline-block;background:var(--danger);color:#fff;
  padding:1px 6px;border-radius:8px;font-size:.7em;margin-left:4px;vertical-align:middle;
  font-weight:700;
}}
.badge.v5{{background:var(--accent)}}

/* ── NOTE ── */
.note{{
  background:#FFFBEB;border-left:4px solid var(--warn);
  padding:12px 16px;border-radius:8px;
  font-size:.86em;color:#78350F;line-height:1.75;
}}
.note strong{{color:#92400E}}

/* ── アップロードエリア ── */
.drop-zone{{
  border:2.5px dashed var(--primary);border-radius:12px;
  padding:40px 24px;text-align:center;cursor:pointer;
  transition:background .3s,border-color .3s,transform .2s;
  position:relative;overflow:hidden;
}}
.drop-zone::before{{
  content:'';position:absolute;inset:0;
  background:radial-gradient(ellipse at 50% 0%,rgba(37,99,235,.08) 0%,transparent 70%);
  opacity:0;transition:opacity .3s;
}}
.drop-zone:hover::before,.drop-zone.over::before{{opacity:1}}
.drop-zone:hover,.drop-zone.over{{
  background:var(--primary-light);border-color:var(--primary-dark);
  transform:translateY(-2px);
}}
.drop-icon{{font-size:2.8em;margin-bottom:10px;animation:bounce 2.5s infinite}}
.drop-text{{color:var(--gray-600);font-size:.95em;line-height:1.6}}
.drop-sub{{color:var(--gray-200);margin:8px 0;font-size:.85em}}
input[type=file]{{display:none}}
.pick-btn{{
  background:linear-gradient(135deg,var(--primary),var(--accent));
  color:#fff;padding:9px 26px;border:none;border-radius:22px;
  cursor:pointer;font-size:.9em;font-weight:600;
  box-shadow:0 4px 12px rgba(37,99,235,.3);transition:transform .2s,box-shadow .2s;
  display:inline-block;
}}
.pick-btn:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(37,99,235,.45)}}
.fname{{
  margin-top:12px;color:var(--primary-dark);font-weight:600;font-size:.9em;
  display:flex;align-items:center;justify-content:center;gap:6px;
  animation:fadeIn .4s ease;
}}

/* ── 生成ボタン ── */
.go-btn{{
  width:100%;background:linear-gradient(135deg,var(--primary),var(--accent));
  color:#fff;padding:14px;border:none;border-radius:22px;
  font-size:1.05em;font-weight:700;cursor:pointer;margin-top:16px;
  transition:transform .25s,box-shadow .25s;
  box-shadow:0 6px 20px rgba(37,99,235,.35);
  position:relative;overflow:hidden;
}}
.go-btn::after{{
  content:'';position:absolute;inset:0;
  background:linear-gradient(90deg,transparent,rgba(255,255,255,.25),transparent);
  transform:translateX(-100%);transition:transform .5s;
}}
.go-btn:hover:not(:disabled)::after{{transform:translateX(100%)}}
.go-btn:hover:not(:disabled){{transform:translateY(-3px);box-shadow:0 10px 28px rgba(37,99,235,.5)}}
.go-btn:disabled{{background:var(--gray-200);color:var(--gray-600);cursor:not-allowed;box-shadow:none}}

/* ── プログレス ── */
.spin-wrap{{display:none;text-align:center;padding:24px 0}}
.progress-bar-wrap{{
  background:var(--gray-200);border-radius:99px;height:8px;
  margin:16px 0 10px;overflow:hidden;
}}
.progress-bar{{
  height:100%;border-radius:99px;width:0%;
  background:linear-gradient(90deg,var(--primary),var(--accent));
  animation:progressAnim 4s ease-in-out infinite;
}}
@keyframes progressAnim{{
  0%{{width:5%}}40%{{width:60%}}70%{{width:75%}}100%{{width:90%}}
}}
.spinner-ring{{
  display:inline-block;width:52px;height:52px;position:relative;margin-bottom:12px;
}}
.spinner-ring div{{
  position:absolute;width:42px;height:42px;margin:4px;
  border:4px solid transparent;border-top-color:var(--primary);
  border-radius:50%;animation:spinRing 1.2s cubic-bezier(.5,0,.5,1) infinite;
}}
.spinner-ring div:nth-child(2){{animation-delay:-0.45s;border-top-color:var(--accent)}}
.spinner-ring div:nth-child(3){{animation-delay:-0.3s;border-top-color:var(--primary);opacity:.6}}
@keyframes spinRing{{to{{transform:rotate(360deg)}}}}
.pmsg{{color:var(--gray-600);font-size:.92em;line-height:1.7}}
.pmsg strong{{color:var(--primary);font-size:1.1em}}

/* ── 成功・エラー ── */
.ok-card{{
  display:none;
  background:linear-gradient(135deg,#ECFDF5,#D1FAE5);
  border:1px solid #A7F3D0;border-radius:12px;
  padding:24px;text-align:center;margin-top:16px;
  animation:popIn .5s ease;
}}
.ok-card p{{color:#064E3B;font-size:1.02em;font-weight:600;margin-bottom:12px}}
.dl-btn{{
  display:inline-flex;align-items:center;gap:8px;
  background:linear-gradient(135deg,var(--success),#10B981);
  color:#fff;padding:11px 32px;text-decoration:none;
  border-radius:22px;font-size:1em;font-weight:700;
  box-shadow:0 4px 14px rgba(5,150,105,.4);
  transition:transform .2s,box-shadow .2s;
}}
.dl-btn:hover{{transform:translateY(-2px);box-shadow:0 8px 20px rgba(5,150,105,.5)}}
.err{{
  display:none;background:#FEF2F2;border:1px solid #FECACA;
  color:#7F1D1D;padding:14px 18px;border-radius:10px;margin-top:16px;
  word-break:break-all;white-space:pre-wrap;font-size:.88em;line-height:1.6;
  animation:fadeIn .4s ease;
}}

/* ── 凡例 ── */
.legend{{display:flex;gap:16px;flex-wrap:wrap;margin-top:6px}}
.legend-item{{display:flex;align-items:center;gap:6px;font-size:.83em;color:var(--gray-600)}}
.sw{{width:16px;height:16px;border-radius:4px;border:1px solid rgba(0,0,0,.1)}}
.c-pink{{background:#FFB6C1}}.c-green{{background:#90EE90}}.c-blue{{background:#BDD7EE}}
.c-a{{background:#DEEAF1}}.c-b{{background:#E2EFDA}}

/* ── フローティングパーティクル ── */
.particles{{position:fixed;inset:0;pointer-events:none;z-index:0;overflow:hidden}}
.particle{{
  position:absolute;width:5px;height:5px;border-radius:50%;
  background:rgba(255,255,255,.15);
  animation:floatUp var(--dur) linear infinite;
  bottom:-10px;left:var(--left);
}}
@keyframes floatUp{{
  0%{{transform:translateY(0) scale(1);opacity:.6}}
  100%{{transform:translateY(-110vh) scale(.3);opacity:0}}
}}

/* ── アニメーション ── */
@keyframes slideDown{{from{{transform:translateY(-30px);opacity:0}}to{{transform:none;opacity:1}}}}
@keyframes fadeUp{{from{{transform:translateY(40px);opacity:0}}to{{transform:none;opacity:1}}}}
@keyframes fadeIn{{from{{opacity:0}}to{{opacity:1}}}}
@keyframes bounce{{0%,100%{{transform:translateY(0)}}50%{{transform:translateY(-8px)}}}}
@keyframes popIn{{0%{{transform:scale(.85);opacity:0}}70%{{transform:scale(1.04)}}100%{{transform:scale(1);opacity:1}}}}

/* ── レスポンシブ ── */
@media(max-width:600px){{
  .section{{padding:20px 18px}}
  .rules-grid{{grid-template-columns:1fr}}
  .logo-text h1{{font-size:1.2em}}
  .logo-icon{{width:48px;height:48px}}
}}
</style>
</head>
<body>

<!-- パーティクル背景 -->
<div class="particles" id="ptc"></div>

<!-- ヘッダー -->
<header class="app-header">
  <div class="logo-block">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
        <path d="M19 3h-1V1h-2v2H8V1H6v2H5C3.89 3 3 3.9 3 5v14c0 1.1.89 2 2 2h14c1.1 0 2-.9
          2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/>
      </svg>
    </div>
    <div class="logo-text">
      <h1>シフト表 自動作成</h1>
      <p>Shift Schedule Generator — AI Optimizer</p>
    </div>
  </div>
  <div class="ver-badge">v5.0&nbsp;LATEST</div>
</header>

<!-- メインカード -->
<div class="card" style="position:relative;z-index:1">

  <!-- 制約セクション -->
  <div class="section">
    <div class="sec-title"><div class="icon">🔒</div>適用される制約・ルール</div>
    <div class="rules-grid">
      <div class="rule-item"><span class="chk">✔</span>ユニットA/B毎日<strong>早出1・遅出1</strong>（兼務職員はどちらかにカウント）</div>
      <div class="rule-item"><span class="chk">✔</span>夜勤1名/日（個人の最少〜最高回数厳守）</div>
      <div class="rule-item"><span class="chk">✔</span>40h→最大5連勤 / 32h・パート→最大4連勤</div>
      <div class="rule-item"><span class="chk">✔</span>夜勤→翌日× / 遅出→翌早禁止</div>
      <div class="rule-item"><span class="chk">✔</span>希望休前日の夜勤禁止、パートに有給自動割当なし</div>
      <div class="rule-item"><span class="chk">✔</span>備考（早出のみ・週N日・夜勤なし等）厳守</div>
      <div class="rule-item"><span class="chk">✔</span>固定公休（曜日指定）対応</div>
      <div class="rule-item"><span class="chk">✔</span>公休数を指定日数に近づける（リーダー以外）<span class="badge">NEW</span></div>
      <div class="rule-item"><span class="chk">✔</span>連続夜勤：○職員のみ緊急時「夜夜××」を許可<span class="badge">NEW</span></div>
      <div class="rule-item"><span class="chk">✔</span>勤務間隔：3〜4日に1回休みを配慮<span class="badge">NEW</span></div>
      <div class="rule-item"><span class="chk">✔</span>同一勤務の連続（早早早等）を回避<span class="badge">NEW</span></div>
      <div class="rule-item"><span class="chk">✔</span>主任は緊急時のみ早出補完（A早/B早で表示）<span class="badge v5">v5</span></div>
      <div class="rule-item"><span class="chk">✔</span>兼務職員の他ユニット勤務をA早/B早/A遅/B遅で表示<span class="badge v5">v5</span></div>
      <div class="rule-item"><span class="chk">✔</span>集計行はCOUNTIF数式（A早/B早/A遅/B遅/夜勤）<span class="badge v5">v5</span></div>
    </div>
  </div>

  <!-- ノート -->
  <div class="section">
    <div class="note">
      <strong>📋 必要なシート：</strong> Staff_Master / Settings / Shift_Requests / Prev_Month / shift_result<br>
      <strong>【ユニット兼務】</strong> Staff_Masterに「ユニット兼務」列を追加し ○ を記入。他ユニット勤務時は <em>A早/B早/A遅/B遅</em> で出力されます。<br>
      <strong>【主任】</strong> ユニット欄が空欄の職員は主任扱い。緊急時のみ早出補完（Excel上で <span style="background:#BDD7EE;padding:0 4px;border-radius:3px">青色</span> 表示）。<br>
      <strong>【集計行】</strong> シフト表下部にCOUNTIF数式でA早/B早/A遅/B遅/夜勤を自動集計します。<br>
      <strong>【ファイル形式】</strong> .xlsx / .xlsm / .xls のいずれも対応しています。
    </div>
  </div>

  <!-- アップロード -->
  <div class="section">
    <div class="sec-title"><div class="icon">📤</div>ファイルアップロード</div>
    <form id="frm">
      <div class="drop-zone" id="drop">
        <div class="drop-icon">📂</div>
        <div class="drop-text">ここにExcelファイルをドラッグ＆ドロップ</div>
        <div class="drop-sub">— または —</div>
        <label for="fi" class="pick-btn">📁 ファイルを選択</label>
        <input type="file" id="fi" accept=".xlsx,.xls,.xlsm">
        <div class="fname" id="fname"></div>
      </div>
      <button type="submit" class="go-btn" id="go">▶&nbsp;&nbsp;シフト表を生成する</button>
    </form>

    <div class="spin-wrap" id="sw">
      <div class="spinner-ring"><div></div><div></div><div></div></div>
      <div class="progress-bar-wrap"><div class="progress-bar" id="pbar"></div></div>
      <p class="pmsg" id="pmsg">最適化中… <strong>0秒</strong> 経過<br>最大5分かかる場合があります。そのままお待ちください。</p>
    </div>
    <div class="ok-card" id="ok">
      <p>✅ シフト表の生成が完了しました！</p>
      <a href="#" id="dl" class="dl-btn">📥 Shift_Result.xlsx をダウンロード</a>
    </div>
    <div class="err" id="er"></div>
  </div>

  <!-- 凡例 -->
  <div class="section">
    <div class="sec-title"><div class="icon">🎨</div>カラー凡例</div>
    <div class="legend">
      <div class="legend-item"><div class="sw c-pink"></div>希望休・有給（希望）</div>
      <div class="legend-item"><div class="sw c-green"></div>勤務指定（Shift_Requests指定）</div>
      <div class="legend-item"><div class="sw c-blue"></div>主任補完（緊急使用）</div>
      <div class="legend-item"><div class="sw c-a"></div>Aユニット</div>
      <div class="legend-item"><div class="sw c-b"></div>Bユニット</div>
    </div>
  </div>

</div><!-- /card -->

<script>
/* ─ パーティクル生成 ─ */
(function(){{
  const c=document.getElementById('ptc');
  for(let i=0;i<22;i++){{
    const p=document.createElement('div');
    p.className='particle';
    p.style.setProperty('--left',Math.random()*100+'%');
    p.style.setProperty('--dur',(6+Math.random()*10).toFixed(1)+'s');
    p.style.animationDelay=(Math.random()*8).toFixed(2)+'s';
    const sz=3+Math.random()*5;
    p.style.width=sz+'px';p.style.height=sz+'px';
    c.appendChild(p);
  }}
}})();

/* ─ ファイル選択 ─ */
const fi=document.getElementById('fi'),fname=document.getElementById('fname'),
      drop=document.getElementById('drop'),frm=document.getElementById('frm'),
      sw=document.getElementById('sw'),ok=document.getElementById('ok'),
      er=document.getElementById('er'),dl=document.getElementById('dl'),
      go=document.getElementById('go'),pmsg=document.getElementById('pmsg');

function setFile(f){{
  if(!f) return;
  const dt=new DataTransfer();dt.items.add(f);fi.files=dt.files;
  fname.innerHTML='<span>📄</span>'+f.name;
}}
fi.onchange=()=>{{ if(fi.files[0]) fname.innerHTML='<span>📄</span>'+fi.files[0].name; }};

['dragenter','dragover','dragleave','drop'].forEach(e=>
  drop.addEventListener(e,ev=>{{ev.preventDefault();ev.stopPropagation();}}));
['dragenter','dragover'].forEach(e=>drop.addEventListener(e,()=>drop.classList.add('over')));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,()=>drop.classList.remove('over')));
drop.addEventListener('drop',e=>{{
  if(e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
}});
drop.addEventListener('click',ev=>{{
  if(ev.target.classList.contains('pick-btn')||ev.target.tagName==='LABEL') return;
  fi.click();
}});

/* ─ タイマー ─ */
let elapsed=0,timer=null;
function startTimer(){{
  elapsed=0;
  timer=setInterval(()=>{{
    elapsed++;
    pmsg.innerHTML='最適化中… <strong>'+elapsed+'秒</strong> 経過<br>最大5分かかる場合があります。そのままお待ちください。';
  }},1000);
}}
function stopTimer(){{if(timer){{clearInterval(timer);timer=null;}}}}

/* ─ フォーム送信 ─ */
frm.onsubmit=async e=>{{
  e.preventDefault();
  if(!fi.files[0]){{alert('ファイルを選択してください');return;}}
  const fd=new FormData();fd.append('file',fi.files[0]);
  sw.style.display='block';ok.style.display='none';er.style.display='none';go.disabled=true;
  startTimer();
  try{{
    const res=await fetch('/generate-shift',{{method:'POST',body:fd}});
    stopTimer();
    if(res.ok){{
      const blob=await res.blob();
      dl.href=URL.createObjectURL(blob);
      dl.download=fi.files[0].name.replace(/[.](xlsx|xlsm|xls)$/i,'')+'_result.xlsx';
      sw.style.display='none';ok.style.display='block';
    }}else{{
      const j=await res.json().catch(()=>({{}}));
      throw new Error(j.detail||'サーバーエラーが発生しました');
    }}
  }}catch(ex){{
    stopTimer();sw.style.display='none';er.style.display='block';
    er.textContent='❌ エラー:\n'+ex.message;
  }}finally{{go.disabled=false;}}
}};
</script>
</body>
</html>"""


# ========================================================
# FastAPI Routes
# ========================================================
@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(content=HTML_CONTENT)

@app.get("/favicon.png")
async def favicon_png():
    if _favicon_path.exists():
        return FileResponse(str(_favicon_path), media_type="image/png")
    raise HTTPException(status_code=404, detail="favicon not found")

@app.get("/favicon.ico")
async def favicon_ico():
    if _favicon_path.exists():
        return FileResponse(str(_favicon_path), media_type="image/png")
    raise HTTPException(status_code=404, detail="favicon not found")

@app.get("/health")
async def health():
    return {"status": "ok", "version": "5.0"}

@app.post("/generate-shift")
async def generate(file: UploadFile = File(...)):
    uid  = str(uuid.uuid4())
    # 入力ファイルの拡張子を保持（xlsm対応）
    orig_name = file.filename or "upload.xlsx"
    ext  = os.path.splitext(orig_name)[1].lower()
    if ext not in [".xlsx", ".xls", ".xlsm"]:
        ext = ".xlsx"
    in_p  = os.path.join(TEMP_DIR, f"in_{uid}{ext}")
    out_p = os.path.join(TEMP_DIR, f"out_{uid}.xlsx")
    try:
        with open(in_p, "wb") as f:
            shutil.copyfileobj(file.file, f)
        (result, staff, shuunin_list, unit_map, cont_map, role_map,
         days_norm, requests, ab_unit_result, shuunin_unit_result,
         kanmu_map) = generate_shift(in_p)
        write_shift_result(
            result, staff, shuunin_list, unit_map, cont_map, role_map,
            days_norm, requests, ab_unit_result, shuunin_unit_result,
            kanmu_map, in_p, out_p)
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
    print(" シフト表自動作成アプリ v5.0")
    print(f" http://localhost:{port}")
    print("=" * 50)
    uvicorn.run("main:app", host=host, port=port, reload=False)
