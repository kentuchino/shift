# シフト表自動作成アプリ v2.3

## ✨ 特徴
- **インストール不要** — ブラウザだけで使えます
- **Excelアップロードだけ** — 最大120秒でシフト表が自動生成
- **クラウド対応** — Renderで無料デプロイ可能
- **USBメモリ対応** — Windows EXEとして持ち運び可能

---

## 🚀 起動方法

### 方法A: クラウド（推奨 / PCにインストール不要）
1. [Render.com](https://render.com) でアカウント作成（無料）
2. このフォルダを GitHub にプッシュ
3. Render で「New > Web Service」→ リポジトリを選択
4. デプロイ完了後、発行された URL をブラウザで開く

### 方法B: Windows EXE（USBメモリ持ち運び）
> **前提**: Python が入っているPCで一度だけビルド作業が必要

```bash
pip install pyinstaller
pyinstaller --onefile --name shift_app_launcher --noconsole launcher.py
# → dist/shift_app_launcher.exe が生成される
```

`shift_app_launcher.exe` を USB に入れて持ち運び。
ダブルクリックでブラウザが自動起動します。

### 方法C: Python直接起動（開発者向け）
```bash
pip install -r requirements.txt
python main.py
# → http://localhost:8000 が自動で開きます
```

---

## 📋 必要なシート（5枚）

| シート名 | 内容 |
|----------|------|
| Staff_Master | 職員情報（ユニット・契約区分・夜勤回数・備考） |
| Settings | 期間・公休数・禁止パターン |
| Shift_Requests | 希望シフト（希望休/有給/指定勤務） |
| Prev_Month | 前月実績（連勤カウント継続） |
| shift_result | 出力テンプレート |

---

## 🔒 実装済み制約（13項目）

| # | 制約 | 内容 |
|---|------|------|
| 1 | ユニット配置 | A/B 毎日早出1・遅出1 |
| 2 | A・B職員 | どちらか一方のユニットのみにカウント |
| 3 | 夜勤人数 | 毎日1名（全体）・個人の最少〜最高回数 |
| 4 | 夜勤後 | 翌日必ず× |
| 5 | 遅出後 | 翌日早出禁止 |
| 6 | 希望休前日 | 夜勤禁止 |
| 7 | 連勤制限 | 40h→5日 / 32h・パート→4日（前月継続） |
| 8 | 希望シフト | 希望休/有給/指定勤務を絶対固定 |
| 9 | 公休数 | Settings指定の最低公休数を確保 |
| 10 | 備考遵守 | 早出のみ/遅出のみ/夜勤なし等 |
| 11 | パート有給 | **自動割り当てなし**（指定時のみ） |
| 12 | 週勤務日数 | パート等の週単位勤務日数（日〜土） |
| 13 | 早遅平準化 | リーダー以外の早出・遅出回数を均等化 |

---

## 🎨 セルの色

- 🩷 **ピンク** — 希望休・有給（Shift_Requestsで指定）
- 💚 **緑** — 勤務指定（Shift_Requestsで指定）

---

## 🛠 技術スタック
- **Backend**: FastAPI (Python 3.11+)
- **Solver**: Google OR-Tools CP-SAT
- **Data**: Pandas, OpenPyXL
- **Frontend**: HTML5 / CSS3 / JavaScript
- **Cloud**: Render (Docker / Python)

---

## 📝 Staff_Master 備考欄の書き方

| 備考の例 | 効果 |
|---------|------|
| `早出のみ` | 遅出・夜勤・日勤を禁止 |
| `週4日勤務` | 週（日〜土）の勤務日数を4日に固定 |
| `早出のみ。週5日勤務。` | 両方を適用 |
| （空欄） | 制限なし（早/遅/夜/日すべて可能） |
