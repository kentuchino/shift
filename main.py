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

app = FastAPI(title="Smart Shift by OR-Tools")
TEMP_DIR = "temp_files"
os.makedirs(TEMP_DIR, exist_ok=True)

WORK_SHIFTS = ["早", "遅", "夜", "日"]
REST_SHIFTS  = ["×", "有", "○"]   # ○=夜勤明け休
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
    nenkyuu = {}
    header_row = None
    for i in range(len(df)):
        v = str(df.iloc[i, 0]).strip()
        if "期間" in v and "開始" in v:
            header_row = i
            break
    if header_row is None:
        raise Exception("Settingsシートに期間ヘッダーが見つかりません")

    for j in range(header_row + 1, len(df)):
        s_val = pd.to_datetime(df.iloc[j, 0], errors="coerce")
        e_val = pd.to_datetime(df.iloc[j, 1], errors="coerce")
        c = str(df.iloc[j, 2]).strip()
        n_str = str(df.iloc[j, 3]).strip()
        # 列4が年休数（新レイアウト）、存在しない場合は空文字
        try:
            nen_str = str(df.iloc[j, 4]).strip() if df.shape[1] > 4 else ""
        except Exception:
            nen_str = ""
        if pd.isna(s_val) and pd.isna(e_val) and c in ["nan", "None", ""]:
            continue
        if pd.notna(s_val):
            start = s_val if start is None else min(start, s_val)
        if pd.notna(e_val):
            end = e_val if end is None else max(end, e_val)
        # 公休数（列3）
        m = re.search(r"\d+", n_str)
        if m and c not in ["nan", "None", ""]:
            num = int(m.group())
            if "40" in c:
                holidays["40h"] = holidays.get("40h", 0) + num
            elif "32" in c:
                holidays["32h"] = holidays.get("32h", 0) + num
            elif "パート" in c:
                holidays["パート"] = holidays.get("パート", 0) + num
        # 年休数（列4）
        nen_m = re.search(r"\d+", nen_str)
        if nen_m and c not in ["nan", "None", ""]:
            nen_num = int(nen_m.group())
            if "40" in c:
                # 同一契約区分で複数行ある場合、最初の非零値を採用
                if nenkyuu.get("40h", 0) == 0:
                    nenkyuu["40h"] = nen_num
            elif "32" in c:
                if nenkyuu.get("32h", 0) == 0:
                    nenkyuu["32h"] = nen_num
            elif "パート" in c:
                if nenkyuu.get("パート", 0) == 0:
                    nenkyuu["パート"] = nen_num

    if start is None or end is None:
        raise Exception("期間が取得できませんでした")

    holidays.setdefault("40h", 9)
    holidays.setdefault("32h", 8)
    holidays.setdefault("パート", 0)
    # 年休デフォルト（Settingsに記載がない場合は従来の2日）
    nenkyuu.setdefault("40h", 2)
    nenkyuu.setdefault("32h", 2)
    nenkyuu.setdefault("パート", 0)

    days = []
    d = start
    while d <= end:
        days.append(d)
        d += timedelta(days=1)
    return days, holidays, nenkyuu


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
            elif "有給" in raw or raw in ("有", "年") or "年休" in raw:
                # 新略称「年」「年休」も有給として認識
                requests[name][date] = ("有", "指定" if is_part else "希望")
            elif "夜勤" in raw or raw == "夜":
                requests[name][date] = ("夜", "指定")
            elif "早出" in raw or raw in ("早", "ハ"):
                # 新略称「ハ」も早出として認識
                requests[name][date] = ("早", "指定")
            elif "遅出" in raw or raw in ("遅", "オ"):
                # 新略称「オ」も遅出として認識
                requests[name][date] = ("遅", "指定")
            elif "日勤" in raw or raw in ("日", "ニ"):
                # 新略称「ニ」も日勤として認識
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
            if "夜勤" in raw or raw == "夜":                          seq.append("夜")
            elif "早出" in raw or raw in ("早", "ハ"):                seq.append("早")
            elif "遅出" in raw or raw in ("遅", "オ"):                seq.append("遅")
            elif "日勤" in raw or raw in ("日", "ニ"):                seq.append("日")
            elif "有給" in raw or raw in ("有", "年") or "年休" in raw: seq.append("有")
            else:                                                      seq.append("×")
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
# INFEASIBLE 診断ヘルパー
# ========================================================
def _diagnose_infeasible(staff, shuunin_list, requests, days_norm, N,
                         allowed_shifts_map, fixed_holiday_map, holiday_limits,
                         cont_map, nmin_map, nmax_map, prev_month, weekly_work_days,
                         unit_map=None, ab_staff_set=None):

    msgs = []
    seen = set()

    def add_msg(m):
        if m not in seen:
            seen.add(m)
            msgs.append(m)

    SHIFT_NAME = {
        "早": "早出(ハ)", "遅": "遅出(オ)", "日": "日勤(ニ)",
        "夜": "夜勤", "×": "休み(×)", "有": "年休(有)", "○": "明け(○)"
    }
    # 「指定」で許可されるシフト（備考制限外でも可）
    ALLOWED_SPECIFIED = {"早", "遅", "夜"}

    # ── 1. 備考制限と希望シフトの矛盾（一般・パートスタッフのみ）──
    # 「指定」の場合は早出・遅出・夜勤のいずれかなら備考制限を超えてもOK
    # 主任はCheck 2 で専用処理するためここでは除外
    for s in staff:
        allowed = allowed_shifts_map.get(s)
        if allowed is None:
            continue
        forbidden_reqs = []
        for date_obj, (sh_type, req_type) in requests.get(s, {}).items():
            if sh_type not in ["早","遅","日","夜"]:
                continue
            is_in_allowed = sh_type in allowed
            is_allowed_specified = (req_type == "指定" and sh_type in ALLOWED_SPECIFIED)
            if not is_in_allowed and not is_allowed_specified and req_type == "指定":
                forbidden_reqs.append(
                    f"{date_obj.strftime('%m/%d')}({SHIFT_NAME.get(sh_type,sh_type)})指定")
        if forbidden_reqs:
            allowed_names = ",".join(SHIFT_NAME.get(a, a) for a in sorted(allowed))
            days_str = ", ".join(forbidden_reqs)
            add_msg(
                f"致命的エラー: {s}さんの 希望シフト・希望休の制約条件が矛盾しているか、データが不正です。\n"
                f"  備考制限: 許可勤務は [{allowed_names}] のみ（指定で 早出・遅出・夜勤 は可）\n"
                f"  矛盾する指定: {days_str}\n"
                f"  → Shift_Requestsシートで該当日の指定を 早出(ハ)・遅出(オ)・夜勤・休み(×) に変更してください。"
            )

    # ── 2. 主任への不正シフト指定（遅・夜の指定はOK、日・有・○は不可）──
    for s in shuunin_list:
        bad_reqs = []
        for date_obj, (sh_type, req_type) in requests.get(s, {}).items():
            if sh_type in ["遅","夜"] and req_type == "指定":
                continue  # 遅出・夜勤の指定はOK
            if sh_type in ["日","有","○"] and req_type == "指定":
                bad_reqs.append(
                    f"{date_obj.strftime('%m/%d')}({SHIFT_NAME.get(sh_type,sh_type)})指定")
        if bad_reqs:
            days_str = ", ".join(bad_reqs)
            add_msg(
                f"致命的エラー: {s}さん（主任）への指定が矛盾しています。\n"
                f"  主任は 早出(ハ)・遅出(オ)・夜勤・休み(×) のみ指定可能です。\n"
                f"  矛盾する指定: {days_str}\n"
                f"  → Shift_Requestsシートで該当日を早出(ハ)・遅出(オ)・夜勤・休み(×)に変更してください。"
            )

    # ── 3. 公休数との矛盾（希望休の数が設定公休数を超える）──
    for s in staff:
        hol_limit = holiday_limits.get(cont_map.get(s, "40h"), 0)
        if hol_limit == 0:
            continue
        hope_off_days = sum(
            1 for _date_obj, (sh_type, req_type) in requests.get(s, {}).items()
            if sh_type in ["×", "有"] and req_type == "希望"
        )
        if hope_off_days > hol_limit:
            add_msg(
                f"致命的エラー: {s}さん の 希望休の数({hope_off_days}日) が "
                f"設定公休数({hol_limit}日) を超えているため、スケジュールを確定できません。\n"
                f"  → 希望休を{hol_limit}日以内に絞るか、公休数設定を見直してください。"
            )

    # ── 4. 夜勤設定との矛盾（夜勤上限と公休数のバランス）──
    for s in staff:
        hol = holiday_limits.get(cont_map.get(s, "40h"), 0)
        nmin = nmin_map.get(s, 0)
        nmax = nmax_map.get(s, 0)
        prev_seq = prev_month.get(s, [])
        prev_night = (bool(prev_seq) and prev_seq[-1] == "夜")
        max_maru = nmax + (1 if prev_night else 0)
        if hol > 0 and max_maru > hol:
            add_msg(
                f"致命的エラー: {s}さん の 夜勤上限回数と公休数のバランスが取れていません"
                f"（夜勤の翌日は休みになるため、勤務できる日数が足りません）。\n"
                f"  公休{hol}日 に対し 夜勤上限{nmax}回"
                + ("＋前月末夜勤(+1)" if prev_night else "")
                + f" = ○{max_maru}日必要ですが、×が{max_maru - hol}日不足します。\n"
                f"  → 夜勤上限を{hol - (1 if prev_night else 0)}以下にするか、"
                f"公休数を{max_maru}以上に増やしてください。"
            )
        if nmin > N:
            add_msg(
                f"致命的エラー: {s}さんの 夜勤最少数({nmin}回) が 対象日数({N}日) を超えています。\n"
                f"  → 夜勤最少数を{N}以下に設定してください。"
            )

    # ── 5. 夜勤可能スタッフ全体の夜勤数チェック ──
    night_capable = [s for s in staff if nmax_map.get(s, 0) > 0]
    total_nmax = sum(nmax_map[s] for s in night_capable)
    total_nmin = sum(nmin_map[s] for s in night_capable)
    if total_nmin > N:
        names = ", ".join(f"{s}({nmin_map[s]}回)" for s in night_capable)
        add_msg(
            f"致命的エラー: 夜勤可能スタッフ全員の夜勤最少数合計({total_nmin}回) が "
            f"対象日数({N}日) を超えています。\n"
            f"  [{names}]\n"
            f"  → 夜勤最少数の合計が{N}以下になるよう各職員の設定を見直してください。"
        )
    if total_nmax < N:
        names = ", ".join(f"{s}({nmax_map[s]}回)" for s in night_capable)
        add_msg(
            f"致命的エラー: 夜勤可能スタッフ全員の夜勤最高数合計({total_nmax}回) が "
            f"対象日数({N}日) より少ないです。\n"
            f"  [{names}]\n"
            f"  → 夜勤最高数の合計が{N}以上になるよう各職員の設定を見直してください。"
        )

    # ── 6. 人数不足チェック（各日の早出カバレッジ簡易確認）──
    if unit_map is not None and days_norm:
        _ab_set = ab_staff_set or set()
        for d, dn in enumerate(days_norm):
            date_str = f"{dn.month}月{dn.day}日"
            a_avail = []
            b_avail = []
            for s in staff:
                req = requests.get(s, {}).get(dn)
                fixed_off = (dn.weekday() in (fixed_holiday_map.get(s) or set()))
                is_req_off = (req and req[0] in ["×", "有"] and req[1] == "指定")
                if fixed_off or is_req_off:
                    continue
                u = unit_map.get(s, "")
                is_ab = (s in _ab_set)
                if u == "A" or is_ab:
                    a_avail.append(s)
                if u == "B" or is_ab:
                    b_avail.append(s)
            if len(a_avail) < 1:
                add_msg(
                    f"致命的エラー: {date_str} の Aユニット・早出 の必要人数(1名)に対して、"
                    f"出勤可能なスタッフが不足しています。\n"
                    f"  → {date_str}前後の希望休・固定休を見直してください。"
                )
            if len(b_avail) < 1:
                add_msg(
                    f"致命的エラー: {date_str} の Bユニット・早出 の必要人数(1名)に対して、"
                    f"出勤可能なスタッフが不足しています。\n"
                    f"  → {date_str}前後の希望休・固定休を見直してください。"
                )

    # ── 7. パート職員の年休希望チェック ──
    for s in [s2 for s2 in staff if cont_map.get(s2) == "パート"]:
        warn_days = []
        for date_obj, (sh_type, req_type) in requests.get(s, {}).items():
            if sh_type == "有" and req_type == "希望":
                warn_days.append(date_obj.strftime("%m/%d"))
        if warn_days:
            add_msg(
                f"警告: {s}さん（パート）の年休(有)希望は無視されます。\n"
                f"  対象日: {', '.join(warn_days)}\n"
                f"  → '指定'に変更するか、空白にしてください。"
            )

    return msgs
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
    days, holiday_limits, nenkyuu_limits = load_settings(settings_df)
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

    # 備考の曜日別シフト指定を解析するヘルパー
    # 例: "日曜：夜勤か×" / "水曜：早出" / "金曜：×か遅出"
    SHIFT_KEYWORD_MAP = {
        "早出": "早", "早": "早", "ハ": "早",
        "遅出": "遅", "遅": "遅", "オ": "遅",
        "夜勤": "夜", "夜": "夜",
        "日勤": "日", "日": "日", "ニ": "日",
        "×": "×", "休み": "×",
        "有給": "有", "年休": "有", "有": "有",
        "○": "○",
    }

    def parse_weekday_note(note):
        """
        備考文字列から {曜日番号: {許可シフトセット}} を返す。
        例: "日曜：夜勤か×、水曜：早出" → {6: {"夜","×"}, 2: {"早"}}
        """
        wd_map = {}
        # パターン: 「曜日名:シフト1かシフト2...」
        for seg in re.split(r"[,、。\n]+", note):
            seg = seg.strip()
            m = re.match(
                r"([月火水木金土日](?:曜(?:日)?)?)[:：](.+)", seg)
            if not m:
                continue
            wd_str = m.group(1)
            shifts_str = m.group(2)
            # 曜日番号
            base = wd_str.replace("曜日","").replace("曜","")
            wd_num = WEEKDAY_MAP.get(base) if base in WEEKDAY_MAP else WEEKDAY_MAP.get(wd_str)
            if wd_num is None:
                continue
            # シフト解析
            allowed_set = set()
            for token in re.split(r"[かまたは/／・\s]+", shifts_str):
                token = token.strip()
                for kw, sh in SHIFT_KEYWORD_MAP.items():
                    if kw in token:
                        allowed_set.add(sh)
                        break
            if allowed_set:
                wd_map[wd_num] = allowed_set
        return wd_map

    # 職員ごとの備考から曜日シフト制約マップを作成
    weekday_shift_map = {}  # {staff_name: {weekday_int: set_of_allowed_shifts}}

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

        # 曜日別シフト制約
        wd_sh = parse_weekday_note(note)
        if wd_sh:
            weekday_shift_map[s] = wd_sh

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

    # ── SettingsのD/E列1-2行目から指定曜日日勤配置を読み込む ──
    # D1/E1: ヘッダー（「指定曜日日勤配置」など）、D2/E2: 曜日名
    nikkin_weekdays = []
    try:
        for col_idx in [3, 4]:
            if settings_df.shape[1] > col_idx and settings_df.shape[0] > 1:
                h_val = str(settings_df.iloc[0, col_idx]).strip()
                d_val = str(settings_df.iloc[1, col_idx]).strip()
                if any(kw in h_val for kw in ["指定", "日勤", "配置"]):
                    for wd_s in re.split(r"[,、・\s/／]+", d_val):
                        wd_s = wd_s.strip()
                        if wd_s in WEEKDAY_MAP:
                            wd_n = WEEKDAY_MAP[wd_s]
                            if wd_n not in nikkin_weekdays:
                                nikkin_weekdays.append(wd_n)
    except Exception:
        pass

    # nikkin日（指定曜日日勤配置が適用されるday index集合）
    nikkin_days_idx = {d for d, dn in enumerate(days_norm) if dn.weekday() in nikkin_weekdays}

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

    # ── 制約3: 前月最終日が夜勤 → 1日目は○か有（年休）のみ──
    # ○が夜勤数を超える場合は有給（年休）を使用（月2日まで）
    for s in staff:
        if prev_month.get(s, []) and prev_month[s][-1] == "夜":
            # 1日目は勤務(早遅日夜)・×は禁止 → ○か有のみ
            for sh_f in ["早","遅","日","夜","×"]:
                model.Add(x[s,0,sh_f] == 0)
    for s in shuunin_list:
        if prev_month.get(s, []) and prev_month[s][-1] == "夜":
            model.Add(xs[s,0,"○"] == 1)  # 主任は夜勤なしなのでこのケースは実質発生しない

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

        # 夜勤（主任が夜勤指定の日はその主任の変数もカウントに含める）
        shuunin_night_vars_d = [xs[s,d,"夜"] for s in shuunin_list
                                if requests.get(s,{}).get(days_norm[d])
                                and requests[s][days_norm[d]][0] == "夜"
                                and requests[s][days_norm[d]][1] == "指定"]
        model.Add(sum(x[s,d,"夜"] for s in staff) + sum(shuunin_night_vars_d) == 1)

    # ── 制約6: 夜勤回数 ──
    for s in staff:
        nt = sum(x[s,d,"夜"] for d in range(N))
        model.Add(nt >= nmin_map[s])
        model.Add(nt <= nmax_map[s])
    # 主任は原則夜勤なし、ただし「指定」がある場合のみ許可
    for s in shuunin_list:
        for d in range(N):
            req = requests.get(s, {}).get(days_norm[d])
            if req and req[0] == "夜" and req[1] == "指定":
                continue  # 夜勤指定はOK
            model.Add(xs[s,d,"夜"] == 0)

    # ── 制約7: 夜勤→翌日は○（夜勤明け休）──
    # (penalty_terms はソフト制約で使用するため、ここで初期化しておく)
    penalty_terms = []
    cn_vars = {}
    for s in staff:
        can_consec = (consec_night_map.get(s, "×") == "○")
        for d in range(N - 1):
            if can_consec:
                # 通常勤務・有給は翌日禁止（○か夜勤のみ）
                for sh in ["早","遅","日","有","×"]:
                    model.Add(x[s,d+1,sh] == 0).OnlyEnforceIf(x[s,d,"夜"])
                cn = model.NewBoolVar(f"cn_{s}_{d}")
                cn_vars[s,d] = cn
                model.AddBoolAnd([x[s,d,"夜"], x[s,d+1,"夜"]]).OnlyEnforceIf(cn)
                model.AddBoolOr([x[s,d,"夜"].Not(), x[s,d+1,"夜"].Not()]).OnlyEnforceIf(cn.Not())
                if d + 3 < N:
                    model.Add(x[s,d+2,"○"] == 1).OnlyEnforceIf(cn)
                    # d+3: 日勤は絶対禁止。早/遅/夜は「どうしても必要な場合」のみ
                    # ソフト制約（超高ペナルティ）で緊急時のみ許容
                    model.Add(x[s,d+3,"日"] == 0).OnlyEnforceIf(cn)
                    for sh_w in ["早","遅","夜"]:
                        _v3 = model.NewBoolVar(f"cd3_{s}_{d}_{sh_w}")
                        model.AddBoolAnd([cn, x[s,d+3,sh_w]]).OnlyEnforceIf(_v3)
                        model.AddBoolOr([cn.Not(), x[s,d+3,sh_w].Not()]).OnlyEnforceIf(_v3.Not())
                        penalty_terms.append((_v3, 800))  # 非常に高いペナルティ
                elif d + 2 < N:
                    model.Add(x[s,d+2,"○"] == 1).OnlyEnforceIf(cn)
                if d + 2 < N:
                    model.Add(x[s,d,"夜"] + x[s,d+1,"夜"] + x[s,d+2,"夜"] <= 2)
            else:
                # 通常は夜勤翌日=○（夜勤明け休）、または有給（出勤扱い）
                # ○が優先だが、公休数が目標を超える場合は有給を使用
                for sh_forbidden in ["早","遅","日","夜","×"]:
                    model.Add(x[s,d+1,sh_forbidden] == 0).OnlyEnforceIf(x[s,d,"夜"])
                # 翌日は○か有のみ（one-shift制約により自動保証）

    for s in shuunin_list:
        for d in range(N - 1):
            model.Add(xs[s,d+1,"○"] == 1).OnlyEnforceIf(xs[s,d,"夜"])

    # ── 制約7-2: ○は必ず前日夜勤の場合のみ発生（夜勤明け以外に○禁止）──
    for s in staff:
        for d in range(N):
            if d == 0:
                prev_seq = prev_month.get(s, [])
                if not (prev_seq and prev_seq[-1] == "夜"):
                    model.Add(x[s, 0, "○"] == 0)
            else:
                model.Add(x[s, d, "○"] == 0).OnlyEnforceIf(x[s, d-1, "夜"].Not())

    # ── 制約7-3: 主任の○は前日夜勤のみ（夜勤なしの日は○禁止）──
    for s in shuunin_list:
        for d in range(N):
            if d == 0:
                prev_seq = prev_month.get(s, [])
                if not (prev_seq and prev_seq[-1] == "夜"):
                    model.Add(xs[s, 0, "○"] == 0)
            else:
                model.Add(xs[s, d, "○"] == 0).OnlyEnforceIf(xs[s, d-1, "夜"].Not())

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

    # ── 制約11: 公休数（×+○）を指定日数に厳密に設定（主任除く）──
    # ○は夜勤明け休（夜勤回数と連動）、×は通常公休
    for s in staff:
        min_hol = holiday_limits.get(cont_map[s], 8)
        if min_hol > 0:
            total_off = (sum(x[s,d,"×"] for d in range(N)) +
                         sum(x[s,d,"○"] for d in range(N)))
            model.Add(total_off >= min_hol)
            model.Add(total_off <= min_hol)  # 多くても少なくてもダメ（等式）

    # ── 制約12: 備考による勤務制限（月間全般制限 + 曜日別制限）──
    # Shift_Requestsの「指定」は備考制限より優先される
    for s in all_staff_names:
        allowed = allowed_shifts_map.get(s)
        var_d = xs if s in shuunin_list else x
        # 月間全般の勤務制限
        if allowed is not None:
            forbidden = set(WORK_SHIFTS) - allowed
            for d in range(N):
                req = requests.get(s, {}).get(days_norm[d])
                for sh in forbidden:
                    # Shift_Requestsの「指定」が備考制限より優先
                    if req and req[0] == sh and req[1] == "指定":
                        continue
                    model.Add(var_d[s,d,sh] == 0)
        # 曜日別の勤務制限（備考に「日曜：夜勤か×」等の記述がある場合）
        wd_sh_rules = weekday_shift_map.get(s, {})
        if wd_sh_rules:
            for d, dn in enumerate(days_norm):
                wd = dn.weekday()
                if wd not in wd_sh_rules:
                    continue
                req = requests.get(s, {}).get(dn)
                # Shift_Requestsの「指定」が曜日制限より優先
                if req and req[1] == "指定":
                    continue
                wd_allowed = wd_sh_rules[wd]
                all_possible = set(ALL_SHIFTS)
                wd_forbidden = all_possible - wd_allowed
                for sh in wd_forbidden:
                    if sh in ("○",):
                        # ○は夜勤翌日のみ（制約7-2で管理）なのでスキップ
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

    # ── 制約13-2: 一般職員の有給は夜勤翌日または明示的な希望のみ ──
    # パートは制約13で対応済み。リーダー・一般職員の有給を制限。
    for s in staff:
        if s in part_staff:
            continue  # パートは制約13で対応
        for d in range(N):
            req = requests.get(s, {}).get(days_norm[d])
            if req and req[0] == "有" and req[1] in ("指定", "希望"):
                continue  # 明示的な希望・指定は許可
            # 有給は夜勤翌日または前月末夜勤の翌日（1日目）のみ許可
            if d == 0:
                prev_seq = prev_month.get(s, [])
                if not (prev_seq and prev_seq[-1] == "夜"):
                    model.Add(x[s, 0, "有"] == 0)
                # 前月末夜勤の場合は1日目は○か有のどちらでも可（制約3で早遅日夜×禁止済み）
            else:
                # 前日が夜勤でない場合は有給禁止
                model.Add(x[s, d, "有"] == 0).OnlyEnforceIf(x[s, d-1, "夜"].Not())

    # ── 制約13-3: 一般職員の年休（有給）は Settingsの年休数（日数）を必ず取得 ──
    # 年休は「出勤扱い」（連勤カウントに含まれる）のため公休数に含まれない
    # Settings で指定した枚数は「一ヶ月に必ず消化しなければならない最低年休数」
    for s in staff:
        if s in part_staff:
            continue
        nen_limit = nenkyuu_limits.get(cont_map.get(s, "40h"), 2)
        if nen_limit > 0:
            total_nenkyuu = sum(x[s,d,"有"] for d in range(N))
            model.Add(total_nenkyuu >= nen_limit)
            # 上限も設ける（=ちょうど nen_limit 枚）
            model.Add(total_nenkyuu <= nen_limit)

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

    # ── 制約15: 主任のシフト制限 ──
    # 通常は早出か×のみ。ただし指定で遅出・夜勤は可。○は前日夜勤時のみ（制約7-3）。
    # nikkin_weekday の日は日勤も許容（指定曜日日勤配置のフォールバック）。
    for s in shuunin_list:
        for d in range(N):
            req = requests.get(s, {}).get(days_norm[d])
            for sh in ["遅","夜","日","有"]:
                # 指定での遅出・夜勤はOK
                if sh in ["遅","夜"] and req and req[0] == sh and req[1] == "指定":
                    continue
                # nikkin_weekday では日勤も許容（フォールバック）
                if sh == "日" and d in nikkin_days_idx:
                    continue
                model.Add(xs[s,d,sh] == 0)
            # ○は制約7-3で管理（前日夜勤のみ許可）

    # ======================================================
    # ソフト制約 & 目的関数
    # ======================================================
    # (penalty_terms は制約7内で初期化済み)

    # ── 制約16: 指定曜日日勤配置（Settings D/E列）──
    # Shift_Requests に日勤指定がある職員とは「別」に、1名日勤を追加配置する
    for d, dn in enumerate(days_norm):
        if dn.weekday() not in nikkin_weekdays:
            continue
        # Shift_Requestsで既に日勤「指定」の職員
        req_nikkin = {s for s in staff
                      if requests.get(s, {}).get(dn)
                      and requests[s][dn][0] == "日"
                      and requests[s][dn][1] == "指定"}
        # 追加日勤対象: req_nikkin 以外の一般職員 + 主任（フォールバック）
        extra_staff_vars   = [x[s,d,"日"]  for s in staff       if s not in req_nikkin]
        extra_shuunin_vars = [xs[s,d,"日"] for s in shuunin_list]
        # ハード: 一般職員か主任のどちらかが1名以上日勤
        model.Add(sum(extra_staff_vars) + sum(extra_shuunin_vars) >= 1)

    # ── ソフト1: 主任使用ペナルティ ──
    for s in shuunin_list:
        for d in range(N):
            penalty_terms.append((xs[s,d,"早"], 200))

    # ── ソフト: 指定曜日日勤配置での主任使用ペナルティ ──
    for d, dn in enumerate(days_norm):
        if dn.weekday() not in nikkin_weekdays:
            continue
        for s in shuunin_list:
            penalty_terms.append((xs[s,d,"日"], 300))

    # ── ソフト2: 連続夜勤使用ペナルティ ──
    for (s, d), cn in cn_vars.items():
        penalty_terms.append((cn, 30))

    # ── ソフト3: 削除済み（ハード制約11で公休数を等式指定に変更）──
    # 制約11が sum(×+○) == target_off を強制するため、ソフト制約は不要

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

    # ── ソフト5-NEW: 遅出翌日の日勤を極力避ける ──
    for s in staff:
        for d in range(N - 1):
            late_then_day = model.NewBoolVar(f"ltd_{s}_{d}")
            model.AddBoolAnd([x[s,d,"遅"], x[s,d+1,"日"]]).OnlyEnforceIf(late_then_day)
            model.AddBoolOr([x[s,d,"遅"].Not(), x[s,d+1,"日"].Not()]).OnlyEnforceIf(late_then_day.Not())
            penalty_terms.append((late_then_day, 10))

    # ── ソフト5-NEW2: ×/有の間隔を10日以内（○はカウント外）──
    # 11日以上の連続×/有なし窓にペナルティ
    for s in staff:
        for start in range(N - 10):
            rest_bits = ([x[s,d,"×"] for d in range(start, start + 11)] +
                         [x[s,d,"有"] for d in range(start, start + 11)])
            gap_viol = model.NewBoolVar(f"gv_{s}_{start}")
            model.Add(sum(rest_bits) == 0).OnlyEnforceIf(gap_viol)
            model.Add(sum(rest_bits) >= 1).OnlyEnforceIf(gap_viol.Not())
            penalty_terms.append((gap_viol, 25))

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

    # ── ソフト7: 夜勤翌日の有給使用にペナルティ（○を優先）──
    # 有給は公休数超過時の調整専用。基本は○を使う。
    for s in [s for s in staff if s not in part_staff]:
        prev_seq_s = prev_month.get(s, [])
        for d in range(N):
            if d == 0:
                # 前月末が夜勤で今月1日が有給の場合
                if prev_seq_s and prev_seq_s[-1] == "夜":
                    penalty_terms.append((x[s, 0, "有"], 15))
            else:
                nty = model.NewBoolVar(f"nty_{s}_{d}")
                model.AddBoolAnd([x[s,d-1,"夜"], x[s,d,"有"]]).OnlyEnforceIf(nty)
                model.AddBoolOr([x[s,d-1,"夜"].Not(), x[s,d,"有"].Not()]).OnlyEnforceIf(nty.Not())
                penalty_terms.append((nty, 15))

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
        # ── INFEASIBLE診断: 誰の何が原因かを特定 ──
        diag_msgs = _diagnose_infeasible(
            staff, shuunin_list, requests, days_norm, N,
            allowed_shifts_map, fixed_holiday_map, holiday_limits,
            cont_map, nmin_map, nmax_map, prev_month, weekly_work_days,
            unit_map=unit_map, ab_staff_set=ab_staff_set
        )
        if diag_msgs:
            raise Exception("\n".join(diag_msgs))
        raise Exception(
            "致命的エラー: 条件を満たすシフト表が見つかりませんでした。\n"
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
            days_norm, requests, ab_unit_result, shuunin_unit_result, kanmu_map, prev_month)


# ========================================================
# Excel 書き出し
# ========================================================
def write_shift_result(result, staff, shuunin_list, unit_map, cont_map, role_map,
                       days_norm, requests, ab_unit_result, shuunin_unit_result,
                       kanmu_map, input_path, output_path, prev_month=None):

    from openpyxl import Workbook
    from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.worksheet.page import PageMargins

    if prev_month is None:
        prev_month = {}

    # ── Settings!B5（期間終了日）を読み込んで太罫線列を決定 ──
    period_end_col_offset = None  # 0-based: day index of period end
    try:
        from openpyxl import load_workbook as _lw
        _wb = _lw(input_path, keep_vba=True, data_only=True)
        _ws = _wb['Settings']
        _b5 = _ws['B5'].value
        if _b5 is not None:
            import pandas as _pd
            _end_date = _pd.to_datetime(_b5).to_pydatetime().replace(
                tzinfo=None, hour=0, minute=0, second=0, microsecond=0)
            for _i, _d in enumerate(days_norm):
                if _d.date() == _end_date.date():
                    period_end_col_offset = _i  # 0-based index
                    break
    except Exception:
        pass

    # ── 後処理なし ──
    result_mod = {}
    all_disp_tmp = shuunin_list + staff
    for s in all_disp_tmp:
        result_mod[s] = dict(result.get(s, {}))

    # ── Workbook 作成 ──
    wb = Workbook()
    ws = wb.active
    ws.title = "shift_result"

    N = len(days_norm)
    weekday_ja = ["月", "火", "水", "木", "金", "土", "日"]
    DATE_START_COL = 2   # 列B = 1日目
    SUMMARY_COL    = DATE_START_COL + N   # 集計列開始

    # 集計列 (10列)
    # SUMMARY_COL+0=早出, +1=日勤, +2=遅出, +3=夜勤, +4=計, +5=○, +6=×, +7=計, +8=年休, +9=合計
    SUMM_ABBR  = ["ハ", "ニ", "オ", "夜勤", "", "", "", "", "年", ""]
    SUMM_FULL  = ["早出", "日勤", "遅出", "夜勤", "計", "○", "×", "計", "年休", "合計"]
    NUM_SUMM   = len(SUMM_FULL)

    # ── 罫線定義 ──
    thin   = Side(style="thin")
    medium = Side(style="medium")
    thin_border   = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
    header_border = Border(left=thin,   right=thin,   top=medium, bottom=medium)

    # ── 職員グループ分け（Staff_Master順を維持）──
    all_staff_ordered = shuunin_list + staff
    group1 = [s for s in all_staff_ordered if unit_map.get(s, "") != "B"]  # 非B(主任+A)
    group2 = [s for s in all_staff_ordered if unit_map.get(s, "") == "B"]  # Bユニット

    STAFF_START_ROW = 5   # 行5から職員データ
    last_group1_row  = STAFF_START_ROW + len(group1) - 1
    first_group2_row = STAFF_START_ROW + len(group1)   # グループ1直後（区切りなし）
    LAST_STAFF_ROW   = STAFF_START_ROW + len(group1) + len(group2) - 1

    # 日別集計行（職員行の2行下から）
    DAILY_ROW_BASE = LAST_STAFF_ROW + 2

    # ── ユニット付きシフト文字列ヘルパー（v5.4: 略称変換付き）──
    SHIFT_ABBR = {"早": "ハ", "遅": "オ", "日": "ニ", "有": "年"}

    def display_val(s, d):
        sh = result_mod[s].get(d, "×")
        # 略称変換（夜/×/○はそのまま）
        if sh == "日":
            return "ニ"
        if sh == "有":
            return "年"
        if sh not in ("早", "遅"):
            return sh  # 夜/×/○
        # 早→ハ, 遅→オ（ユニットプレフィックス付き）
        abbr = SHIFT_ABBR[sh]
        if s in shuunin_list:
            unit = shuunin_unit_result.get(s, {}).get(d)
            return (unit + abbr) if unit else abbr
        elif kanmu_map.get(s, "×") == "○":
            unit = ab_unit_result.get(s, {}).get(d)
            return (unit + abbr) if unit else abbr
        else:
            unit = unit_map.get(s, "")
            return (unit + abbr) if unit in ("A", "B") else abbr

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
            sh   = result.get(s, {}).get(d, "×")
            if sh == "早" and unit:
                return BLUE_FILL
        return None

    # ── 行1: 月ラベル ──
    # Windows互換: %-m月 は Linux のみ動作するため month 属性で代替
    month_label = f"{days_norm[0].month}月" if days_norm else ""
    label_col = DATE_START_COL + N // 2  # 日付範囲の中央付近
    c = ws.cell(1, label_col, month_label)
    c.font = Font(bold=True, size=14)
    c.alignment = Alignment(horizontal="center")

    # ── 行2: 日付番号 ──
    ws.cell(2, 1, "日").alignment = Alignment(horizontal="center")
    ws.cell(2, 1).font = Font(bold=True)
    for i, d in enumerate(days_norm):
        col = DATE_START_COL + i
        c = ws.cell(2, col, d.day)
        c.alignment = Alignment(horizontal="center")
        c.font = Font(bold=True)
        c.border = thin_border

    # ── 行3: 曜日 + 集計略称 ──
    ws.cell(3, 1, "曜日").alignment = Alignment(horizontal="center")
    ws.cell(3, 1).font = Font(bold=True)
    for i, d in enumerate(days_norm):
        col = DATE_START_COL + i
        wd  = weekday_ja[d.weekday()]
        c   = ws.cell(3, col, wd)
        c.alignment = Alignment(horizontal="center")
        c.border = thin_border
        if d.weekday() == 5:   # 土
            c.fill = PatternFill("solid", fgColor="CCE5FF")
        elif d.weekday() == 6: # 日
            c.fill = PatternFill("solid", fgColor="FFCCCC")

    # 集計略称（行3）
    for k, abbr in enumerate(SUMM_ABBR):
        if abbr:
            c = ws.cell(3, SUMMARY_COL + k, abbr)
            c.alignment = Alignment(horizontal="center")
            c.font = Font(bold=True)
            c.fill = YELLOW_FILL
            c.border = thin_border

    # ── 行4: 会議・委員会 + 集計正式名 ──
    c4 = ws.cell(4, 1, "会議・委員会")
    c4.alignment = Alignment(horizontal="center", vertical="center")
    c4.font = Font(bold=True)
    c4.border = thin_border
    # 日付部分（行4）は空セルに細罫線のみ
    for i in range(N):
        ws.cell(4, DATE_START_COL + i).border = thin_border

    # 集計正式名（行4）
    for k, h in enumerate(SUMM_FULL):
        c = ws.cell(4, SUMMARY_COL + k, h)
        c.alignment = Alignment(horizontal="center")
        c.font = Font(bold=True)
        c.fill = YELLOW_FILL
        c.border = header_border

    # ── 職員行を書き込むヘルパー ──
    def write_staff_row(row, s, extra_top=False, extra_bottom=False):
        u = unit_map.get(s, "")
        k = kanmu_map.get(s, "×")
        is_shuunin = (s in shuunin_list)

        # 名前列（列A）
        nc = ws.cell(row, 1, s)
        nc.alignment = Alignment(horizontal="center", vertical="center")
        if is_shuunin:
            nc.fill = BLUE_FILL
            nc.font = Font(bold=True)
        elif u == "A":
            nc.fill = A_UNIT_FILL
        elif u == "B":
            nc.fill = B_UNIT_FILL

        # 日付セル
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
        hc_col = ws.cell(row, SUMMARY_COL,     f'=COUNTIF({rng},"Aハ")+COUNTIF({rng},"Bハ")')   # 早出
        hn_col = ws.cell(row, SUMMARY_COL + 1, f'=COUNTIF({rng},"ニ")')                          # 日勤
        ho_col = ws.cell(row, SUMMARY_COL + 2, f'=COUNTIF({rng},"Aオ")+COUNTIF({rng},"Bオ")')   # 遅出
        hy_col = ws.cell(row, SUMMARY_COL + 3, f'=COUNTIF({rng},"夜")')                          # 夜勤
        k_col  = SUMMARY_COL + 4
        hk     = ws.cell(row, k_col)
        # 計（早出+日勤+遅出+夜勤）
        hk.value = (f'={get_column_letter(SUMMARY_COL)}{row}'
                    f'+{get_column_letter(SUMMARY_COL+1)}{row}'
                    f'+{get_column_letter(SUMMARY_COL+2)}{row}'
                    f'+{get_column_letter(SUMMARY_COL+3)}{row}')
        ws.cell(row, SUMMARY_COL + 5, f'=COUNTIF({rng},"○")')  # ○
        ws.cell(row, SUMMARY_COL + 6, f'=COUNTIF({rng},"×")')  # ×
        # 計（公休: ○+×）
        ws.cell(row, SUMMARY_COL + 7,
                f'={get_column_letter(SUMMARY_COL+5)}{row}'
                f'+{get_column_letter(SUMMARY_COL+6)}{row}')
        ws.cell(row, SUMMARY_COL + 8, f'=COUNTIF({rng},"年")')  # 年休
        # 合計（計+計+年休）
        ws.cell(row, SUMMARY_COL + 9,
                f'={get_column_letter(SUMMARY_COL+4)}{row}'
                f'+{get_column_letter(SUMMARY_COL+7)}{row}'
                f'+{get_column_letter(SUMMARY_COL+8)}{row}')
        for k2 in range(NUM_SUMM):
            c = ws.cell(row, SUMMARY_COL + k2)
            c.alignment = Alignment(horizontal="center")
            c.fill = YELLOW_FILL

        # ── 罫線を行全体に適用 ──
        # top/bottom の調整
        top_side    = medium if extra_top    else thin
        bottom_side = medium if extra_bottom else thin

        # 名前列(A列)
        ws.cell(row, 1).border = Border(
            left=medium, right=thin, top=top_side, bottom=bottom_side)
        # 日付列
        for d in range(N):
            col = DATE_START_COL + d
            ws.cell(row, col).border = Border(
                left=thin, right=thin, top=top_side, bottom=bottom_side)
        # 集計列
        for k2 in range(NUM_SUMM):
            ws.cell(row, SUMMARY_COL + k2).border = Border(
                left=thin, right=thin, top=top_side, bottom=bottom_side)

    # ── グループ1（非B: 主任+Aユニット）を書き込む ──
    for idx, s in enumerate(group1):
        row = STAFF_START_ROW + idx
        is_last  = (idx == len(group1) - 1) and len(group2) > 0
        write_staff_row(row, s, extra_bottom=is_last)

    # ── グループ2（Bユニット）を書き込む ──
    for idx, s in enumerate(group2):
        row = first_group2_row + idx
        is_first = (idx == 0)
        write_staff_row(row, s, extra_top=is_first)

    # ── Settings!B5 の日付列に太い罫線 ──
    if period_end_col_offset is not None:
        period_end_col = DATE_START_COL + period_end_col_offset  # 該当日付列
        next_col       = period_end_col + 1
        # 全行（ヘッダー+職員+集計行）に適用
        apply_rows = list(range(2, LAST_STAFF_ROW + 1)) + list(range(DAILY_ROW_BASE, DAILY_ROW_BASE + 6))
        for row in apply_rows:
            # 期間終了日列: right=medium
            c = ws.cell(row, period_end_col)
            old_b = c.border
            c.border = Border(
                left  = old_b.left  if old_b.left  else thin,
                right = medium,
                top   = old_b.top   if old_b.top   else thin,
                bottom= old_b.bottom if old_b.bottom else thin)
            # 次の列: left=medium
            if next_col <= SUMMARY_COL - 1:  # 集計列の手前まで
                c2 = ws.cell(row, next_col)
                old_b2 = c2.border
                c2.border = Border(
                    left  = medium,
                    right = old_b2.right  if old_b2.right  else thin,
                    top   = old_b2.top    if old_b2.top    else thin,
                    bottom= old_b2.bottom if old_b2.bottom else thin)

    # ── 日別集計行（COUNTIF数式）──
    daily_labels = ["A早出", "B早出", "A遅出", "B遅出", "夜勤", "日勤"]
    daily_codes  = ["Aハ",   "Bハ",   "Aオ",  "Bオ",  "夜",  "ニ"]
    daily_fills  = [A_UNIT_FILL, B_UNIT_FILL, A_UNIT_FILL, B_UNIT_FILL, GRAY_FILL, GRAY_FILL]

    cnt_start_row = STAFF_START_ROW
    cnt_end_row   = LAST_STAFF_ROW

    for k, (lbl, dv, fill) in enumerate(zip(daily_labels, daily_codes, daily_fills)):
        dr = DAILY_ROW_BASE + k
        c = ws.cell(dr, 1, lbl)
        c.fill = fill
        c.alignment = Alignment(horizontal="center")
        c.font = Font(bold=True)
        c.border = Border(left=medium, right=thin, top=thin, bottom=thin)

        for i in range(N):
            col = DATE_START_COL + i
            col_letter = get_column_letter(col)
            cnt_range = f"{col_letter}{cnt_start_row}:{col_letter}{cnt_end_row}"
            formula_val = f'=COUNTIF({cnt_range},"{dv}")'
            dc = ws.cell(dr, col, formula_val)
            dc.alignment = Alignment(horizontal="center")
            dc.border = thin_border

        # 集計列は空白（行ラベルのみ）
        for k2 in range(NUM_SUMM):
            ws.cell(dr, SUMMARY_COL + k2).border = thin_border

    # ── 条件付き書式: 日別集計行の値が0の場合は赤字 ──
    red_font_rule = CellIsRule(
        operator="equal",
        formula=["0"],
        font=Font(color="FF0000", bold=True))
    first_date_letter = get_column_letter(DATE_START_COL)
    last_date_letter  = get_column_letter(DATE_START_COL + N - 1)
    for k in range(len(daily_labels)):
        dr = DAILY_ROW_BASE + k
        range_str = f"{first_date_letter}{dr}:{last_date_letter}{dr}"
        ws.conditional_formatting.add(range_str, red_font_rule)

    # ── 空白行（職員行〜日別集計行の間）の罫線 ──
    sep_row = LAST_STAFF_ROW + 1
    for col in range(1, SUMMARY_COL + NUM_SUMM):
        ws.cell(sep_row, col).border = thin_border

    # ── 行1のヘッダー列に罫線 ──
    ws.cell(1, 1).border = Border(left=medium, right=thin, top=thin, bottom=thin)
    for i in range(N):
        ws.cell(1, DATE_START_COL + i).border = thin_border

    # ── 行2・行3のA列に罫線 ──
    for r in [2, 3, 4]:
        ws.cell(r, 1).border = Border(left=medium, right=thin, top=thin, bottom=thin)

    # ── 列幅・行高さ ──
    ws.column_dimensions["A"].width = 10
    for i in range(N):
        ws.column_dimensions[get_column_letter(DATE_START_COL + i)].width = 5
    for k in range(NUM_SUMM):
        ws.column_dimensions[get_column_letter(SUMMARY_COL + k)].width = 7

    for r in range(1, 5):
        ws.row_dimensions[r].height = 18
    for r in range(STAFF_START_ROW, LAST_STAFF_ROW + 1):
        ws.row_dimensions[r].height = 16
    for k in range(len(daily_labels)):
        ws.row_dimensions[DAILY_ROW_BASE + k].height = 16

    # ── 印刷設定（A4横向き）──
    from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9   # A4
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.7, bottom=0.7, header=0.3, footer=0.3)
    ws.print_title_rows = "1:4"

    wb.save(output_path)


# ========================================================
# Web UI HTML
# ========================================================
_favicon_tag = (f'<link rel="icon" type="image/png" href="data:image/png;base64,{FAVICON_B64}">'
                if FAVICON_B64 else '<link rel="icon" href="/favicon.png">')

HTML_CONTENT = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <title>Smart Shift by OR-Tools</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&family=Noto+Sans+JP:wght@400;700&family=JetBrains+Mono&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        :root {
            --primary: #0066ff;
            --accent: #00ff99;
            --bg-deep: #0a0c10;
            --bg-panel: #161b22;
            --text-main: #e6edf3;
            --text-dim: #8b949e;
            --border: rgba(240, 246, 252, 0.1);
        }

        body { 
            margin: 0; 
            font-family: 'Inter', 'Noto Sans JP', sans-serif; 
            background: var(--bg-deep); 
            color: var(--text-main); 
            display: flex; 
            height: 100vh; 
            overflow: hidden; 
        }

        /* --- サイドパネル --- */
        .panel { 
            width: 280px; 
            background: var(--bg-panel); 
            border-right: 1px solid var(--border); 
            display: flex; 
            flex-direction: column; 
            padding: 20px; 
            flex-shrink: 0; 
            overflow-y: auto; 
        }
        .panel-right { border-right: none; border-left: 1px solid var(--border); width: 320px; }
        
        .header-logo { font-size: 1.1rem; font-weight: 800; color: #fff; margin-bottom: 4px; letter-spacing: 1px; }
        .version-tag { font-size: 0.6rem; color: var(--text-dim); margin-bottom: 24px; }

        .section-label { 
            font-size: 0.75rem; 
            color: var(--accent); 
            font-weight: 700; 
            margin: 20px 0 10px; 
            display: flex; 
            align-items: center; 
            gap: 6px; 
            border-left: 3px solid var(--accent);
            padding-left: 8px;
        }

        /* --- カード表示 --- */
        .info-card { 
            background: rgba(255,255,255,0.03); 
            border-radius: 6px; 
            padding: 12px; 
            margin-bottom: 10px; 
            border: 1px solid var(--border); 
        }
        .info-label { font-size: 0.7rem; color: var(--text-dim); margin-bottom: 4px; display: block; }
        .info-value { font-size: 1rem; font-weight: 600; font-family: 'JetBrains Mono'; }
        
        /* 負荷インジケーター */
        .bar-container { height: 4px; background: rgba(255,255,255,0.1); border-radius: 2px; margin-top: 8px; overflow: hidden; }
        .bar-fill { height: 100%; background: var(--primary); width: 0%; transition: width 0.4s, background 0.4s; }

        /* --- 勤務割付状況（右パネル） --- */
        .rule-grid { display: grid; grid-template-columns: 1fr; gap: 6px; }
        .rule-item { 
            font-size: 0.75rem; padding: 8px; 
            background: rgba(255,255,255,0.02); border-radius: 4px; 
            display: flex; align-items: center; justify-content: space-between;
        }
        .status-light { width: 8px; height: 8px; border-radius: 50%; background: #333; }
        .light-active { background: var(--accent); box-shadow: 0 0 10px var(--accent); animation: blink 0.8s infinite; }

        @keyframes blink { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }

        /* --- 中央 メイン --- */
        .main { flex-grow: 1; display: flex; flex-direction: column; background: radial-gradient(circle at center, #111827 0%, #0a0c10 100%); }
        .top-bar { height: 50px; border-bottom: 1px solid var(--border); display: flex; align-items: center; padding: 0 24px; font-size: 0.8rem; color: var(--text-dim); }

        .workspace { padding: 40px; display: flex; flex-direction: column; align-items: center; gap: 24px; flex-grow: 1; }

        /* ドラッグドロップ */
        .drop-area {
            width: 100%; max-width: 640px; height: 180px; border: 2px dashed #30363d; border-radius: 12px;
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            background: rgba(22, 27, 34, 0.5); transition: 0.2s; cursor: pointer;
        }
        .drop-area:hover, .drop-area.drag-over { border-color: var(--primary); background: rgba(0, 102, 255, 0.05); }

        .log-monitor {
            width: 100%; max-width: 800px; height: 320px; background: #000; border-radius: 8px;
            padding: 16px; font-family: 'JetBrains Mono'; font-size: 0.8rem;
            overflow-y: auto; border: 1px solid #30363d; position: relative;
        }
        .log-row { color: #a3e635; margin-bottom: 3px; line-height: 1.4; }
        .laser-scan { position: absolute; width: 100%; height: 2px; background: rgba(0, 255, 153, 0.3); box-shadow: 0 0 15px var(--accent); display: none; animation: scan 3s infinite linear; }
        @keyframes scan { from { top: 0; } to { top: 100%; } }

        .exec-button {
            width: 100%; max-width: 400px; padding: 16px; border-radius: 8px; border: none;
            background: var(--primary); color: white; font-weight: 700; font-size: 1rem;
            cursor: pointer; display: flex; justify-content: center; align-items: center; gap: 10px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.5);
        }
        .exec-button:hover:not(:disabled) { background: #1a75ff; transform: translateY(-1px); }
        .exec-button:disabled { background: #21262d; color: #484f58; cursor: not-allowed; }

        .loader { width: 18px; height: 18px; border: 2px solid rgba(255,255,255,0.2); border-top-color: #fff; border-radius: 50%; animation: spin 0.8s linear infinite; display: none; }
        @keyframes spin { to { transform: rotate(360deg); } }


        /* --- メインタイトルエリア --- */
        .main-header {
            padding: 30px 40px 10px;
            display: flex;
            flex-direction: column;
            align-items: flex-start;
        }

        .logo-container {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 8px;
            animation: fadeInSlide 1s ease-out;
        }

        .logo-symbol {
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 0 20px rgba(0, 102, 255, 0.4);
        }

        .brand-title {
            font-size: 1.8rem;
            font-weight: 800;
            background: linear-gradient(to right, #fff, var(--text-dim));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.5px;
        }

        .brand-subtitle {
            font-family: 'JetBrains Mono';
            font-size: 0.75rem;
            color: var(--accent);
            text-transform: uppercase;
            letter-spacing: 3px;
            margin-left: 2px;
            opacity: 0.8;
        }

        .header-line {
            width: 100%;
            max-width: 640px;
            height: 1px;
            background: linear-gradient(to right, var(--primary), transparent);
            margin-top: 15px;
            position: relative;
            overflow: hidden;
        }

        .header-line::after {
            content: '';
            position: absolute;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(to right, transparent, var(--accent), transparent);
            animation: lineScan 3s infinite linear;
        }

        @keyframes fadeInSlide {
            from { opacity: 0; transform: translateX(-20px); }
            to { opacity: 1; transform: translateX(0); }
        }

        @keyframes lineScan {
            0% { left: -100%; }
            100% { left: 100%; }
        }
    </style>
</head>
<body>

<aside class="panel">
    

    <div class="section-label">システム正常性</div>
    <div class="info-card">
        <span class="info-label">計算負荷</span>
        <div class="info-value" id="loadStatus">待機中</div>
        <div class="bar-container"><div class="bar-fill" id="loadBar" style="width:2%"></div></div>
    </div>

    <div class="section-label">解析済みデータ</div>
    <div id="fileInspector">
        <div style="font-size:0.75rem; color:var(--text-dim);">ファイルが未投入です</div>
    </div>

    <div class="section-label">エンジン構成</div>
    <div style="font-size:0.7rem; color:var(--text-dim); line-height:2;">
        ・マルチスレッド： 有効<br>
        ・制約モデル： CP-SAT<br>
        ・タイムアウト： 300秒<br>
        ・優先度： 最適解重視
    </div>
</aside>

<main class="main">
   <header class="main-header">
        <div class="logo-container">
            <div class="logo-symbol">
                <i data-lucide="cpu" color="#fff" size="24"></i>
            </div>
            <div>
                <div class="brand-title">Smart Shift <span style="font-weight:300;">by OR-Tools v5.8</span></div>
                <div class="brand-subtitle">The Intelligent Auto-Roster Engine</div>
            </div>
        </div>
        <div class="header-line"></div>
    </header>
    
    <div class="workspace">
        <div id="dropZone" class="drop-area">
            <i data-lucide="file-spreadsheet" size="36" style="margin-bottom:12px; color:var(--primary);"></i>
            <span style="font-weight:700;">Excelファイルをドロップ</span>
            <span id="filePrompt" style="font-size:0.75rem; color:var(--text-dim); margin-top:8px;">(ここにファイルをドラッグしてください)</span>
            <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsm" style="display:none;">
        </div>

        <button id="runBtn" class="exec-button" disabled>
            <div class="loader" id="loader"></div>
            <span id="btnLabel">最適化計算を実行</span>
        </button>

        <div class="log-monitor" id="logMonitor">
            <div class="laser-scan" id="laser"></div>
            <div id="logBody"></div>
        </div>
    </div>
</main>

<aside class="panel panel-right">
    <div class="section-label">最適化メトリクス</div>
    <div class="info-card">
        <span class="info-label">経過時間（ライブ）</span>
        <div class="info-value" id="elapsedTime">0.00秒</div>
    </div>
    
    <div class="info-card">
        <span class="info-label">計算適合率（精度）</span>
        <div class="info-value" id="scoreValue" style="color:var(--accent);">--</div>
        <div class="bar-container"><div class="bar-fill" id="scoreBar"></div></div>
    </div>
    <div class="info-card">
        <span class="info-label">処理時間</span>
        <div class="info-value" id="timeValue">--</div>
    </div>

    <div class="section-label">ハード制約（絶対条件）</div>
    <div class="rule-grid">
        <div class="rule-item"><span>1. Shift_Requestsの希望・指定を厳守</span><div class="status-light" id="L1"></div></div>
        <div class="rule-item"><span>2. 前月最終日が夜勤の場合、当月1日目は「○」</span><div class="status-light" id="L2"></div></div>
        <div class="rule-item"><span>3. Staff_Masterの固定公休欄に従い指定曜日は「×」</span><div class="status-light" id="L3"></div></div>
        <div class="rule-item"><span>4. ユニットA/B 各1名早出・1名遅出、夜勤全体1名</span><div class="status-light" id="L4"></div></div>
        <div class="rule-item"><span>5. Staff_Masterの夜勤最少〜最高を厳守</span><div class="status-light" id="L5"></div></div>
        <div class="rule-item"><span>6. 夜勤翌日は必ず「○」、通常公休「×」との区別</span><div class="status-light" id="L6"></div></div>
        <div class="rule-item"><span>7. 連続夜勤（○可能職員）の翌2日は「○」</span><div class="status-light" id="L7"></div></div>
        <div class="rule-item"><span>8. 遅→翌早は絶対禁止</span><div class="status-light" id="L8"></div></div>
        <div class="rule-item"><span>9. 希望休・有給の前日に夜勤を入れない</span><div class="status-light" id="L9"></div></div>
        <div class="rule-item"><span>10. 40h：最大5連勤、32h・パート：最大4連勤</span><div class="status-light" id="L10"></div></div>
        <div class="rule-item"><span>11. 公休数下限 Settingsの設定値以上の「×」を確保</span><div class="status-light" id="L11"></div></div>
        <div class="rule-item"><span>12. 備考による勤務制限</span><div class="status-light" id="L12"></div></div>
        <div class="rule-item"><span>13. 週単位勤務日数 「週N日勤務」の備考に従いパートの勤務日数を管理</span><div class="status-light" id="L13"></div></div>        
    </div>

    <div class="section-label">ソフト制約（ペナルティ最小化）</div>
    <div class="rule-grid">
        <div class="rule-item"><span>1. 主任の補充は最小限に</span><div class="status-light" id="L14"></div></div>
        <div class="rule-item"><span>2. 連続夜勤は避ける</span><div class="status-light" id="L15"></div></div>
        <div class="rule-item"><span>3. 公休数の目標値近似</span><div class="status-light" id="L16"></div></div>
        <div class="rule-item"><span>4. リーダー以外の早・遅勤務数を均等に</span><div class="status-light" id="L17"></div></div>
        <div class="rule-item"><span>5. 遅出翌日に日勤を極力入れない</span><div class="status-light" id="L18"></div></div>
        <div class="rule-item"><span>6. ×/有なし11日連続をペナルティ化</span><div class="status-light" id="L19"></div></div>
        <div class="rule-item"><span>7. 4日連続勤務を避ける</span><div class="status-light" id="L20"></div></div>
        <div class="rule-item"><span>8. 早/遅の3日連続を避ける</span><div class="status-light" id="L21"></div></div>    
    </div>

    
</aside>

<script>
    lucide.createIcons();
    let targetFile = null;
    let timerInterval = null; // タイマーをグローバルで管理

    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const runBtn = document.getElementById('runBtn');
    const logBody = document.getElementById('logBody');
    const laser = document.getElementById('laser');

    // ドラッグ＆ドロップイベント
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(n => {
        dropZone.addEventListener(n, e => { e.preventDefault(); e.stopPropagation(); });
    });
    ['dragenter', 'dragover'].forEach(n => dropZone.addEventListener(n, () => dropZone.classList.add('drag-over')));
    ['dragleave', 'drop'].forEach(n => dropZone.addEventListener(n, () => dropZone.classList.remove('drag-over')));

    dropZone.addEventListener('drop', e => {
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) handleFile(fileInput.files[0]);
    });

    function handleFile(file) {
        targetFile = file;
        runBtn.disabled = false;
        document.getElementById('filePrompt').textContent = `準備完了: ${file.name}`;
        document.getElementById('filePrompt').style.color = 'var(--accent)';
        
        document.getElementById('fileInspector').innerHTML = `
            <div class="info-card">
                <span class="info-label">対象ファイル名</span>
                <div style="font-size:0.75rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${file.name}</div>
            </div>
            <div class="info-card">
                <span class="info-label">データ整合性</span>
                <div style="font-size:0.75rem; color:var(--accent);">正常（Excel形式）</div>
            </div>
        `;
        addLog(`ファイルを読み込みました: ${file.name}`);
    }

    function addLog(msg, color = 'inherit') {
        const div = document.createElement('div');
        div.className = 'log-row';
        div.style.color = color;
        div.textContent = `> [${new Date().toLocaleTimeString()}] ${msg}`;
        logBody.appendChild(div);
        const monitor = document.getElementById('logMonitor');
        monitor.scrollTop = monitor.scrollHeight;
    }

    runBtn.addEventListener('click', async () => {
        if (!targetFile) return;

        // 初期化
        runBtn.disabled = true;
        document.getElementById('loader').style.display = 'block';
        document.getElementById('btnLabel').textContent = '最適化実行中...';
        laser.style.display = 'block';
        
        // --- 経過時間タイマー開始 ---
        const startTime = performance.now();
        const elapsedDisplay = document.getElementById('elapsedTime');
        const timeValueDisplay = document.getElementById('timeValue');
        
        // 以前のタイマーがあればクリア
        if(timerInterval) clearInterval(timerInterval);
        
        timerInterval = setInterval(() => {
            const now = performance.now();
            const diff = ((now - startTime) / 1000).toFixed(2);
            elapsedDisplay.textContent = diff + '秒';
        }, 50);

        // システム負荷演出
        document.getElementById('loadStatus').textContent = '高負荷';
        document.getElementById('loadBar').style.width = '98%';
        document.getElementById('loadBar').style.background = '#ff4d4d';

        // ルール点灯演出
        const lights = ['L1','L2','L3','L4','L5','L6','L7','L8','L9','L10','L11','L12','L13','L14','L15','L16','L17','L18','L19','L20','L21'].map(id => document.getElementById(id));
        const lightTimer = setInterval(() => {
            lights.forEach(l => l.className = Math.random() > 0.5 ? 'status-light light-active' : 'status-light');
        }, 200);


        addLog("最適化モデルを初期化中...");
        addLog("制約条件のマッピングを開始...");
        
        const fd = new FormData();
        fd.append("file", targetFile);

        try {
            const res = await fetch("/generate-shift", {method: "POST", body: fd});
            if (res.ok) {
                // 完了時の処理
                const finalTime = ((performance.now() - startTime) / 1000).toFixed(2);
                timeValueDisplay.textContent = `${finalTime}秒`;
                document.getElementById('scoreValue').textContent = '100%';
                document.getElementById('scoreBar').style.width = '100%';
                
                addLog("最適解の構築が成功しました。");
                
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url; a.download = "SS_Result.xlsx"; a.click();
            } else {
                // サーバーエラーの詳細を取得してログに表示
                let errDetail = "（詳細不明）";
                try {
                    const errJson = await res.json();
                    if (errJson && errJson.detail) errDetail = errJson.detail;
                } catch(_) {}
                // 改行ごとに分けて赤字で表示
                const errLines = errDetail.split("\n");
                addLog("═══ サーバーエラー発生 ═══", "#ff4d4d");
                errLines.forEach(line => {
                    if (line.trim()) addLog(line, "#ff6b6b");
                });
                addLog("─── エラー終了 ─── データまたは制約条件を見直してください。", "#ff4d4d");
            }
        } catch (e) {
            addLog("ネットワークエラー: サーバーへの接続に失敗しました。", "#ff4d4d");
            addLog(String(e), "#ff6b6b");
        } finally {
            // タイマー停止
            clearInterval(timerInterval);
            clearInterval(lightTimer);
            
            lights.forEach(l => l.className = 'status-light');
            runBtn.disabled = false;
            document.getElementById('loader').style.display = 'none';
            document.getElementById('btnLabel').textContent = '再計算を実行';
            laser.style.display = 'none';
            document.getElementById('loadStatus').textContent = '待機中';
            document.getElementById('loadBar').style.width = '5%';
            document.getElementById('loadBar').style.background = 'var(--primary)';
        }
    });
</script>
</body>
</html>
"""


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
    return {"status": "ok", "version": "5.7"}

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
         kanmu_map, prev_month) = generate_shift(in_p)
        write_shift_result(
            result, staff, shuunin_list, unit_map, cont_map, role_map,
            days_norm, requests, ab_unit_result, shuunin_unit_result,
            kanmu_map, in_p, out_p, prev_month=prev_month)
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
    print(" Smart Shift by OR-Tools")
    print(f" http://localhost:{port}")
    print("=" * 50)
    uvicorn.run("main:app", host=host, port=port, reload=False)
