#!/bin/bash
echo "============================================================"
echo " シフト表自動作成アプリ v3.0 - 起動中..."
echo "============================================================"
echo ""
if ! command -v python3 &>/dev/null; then
    echo "[エラー] python3が見つかりません"
    exit 1
fi
python3 -c "import fastapi" 2>/dev/null || {
    echo "ライブラリをインストールしています..."
    pip3 install -r requirements.txt --quiet
}
echo "ブラウザが自動で開きます。開かない場合は http://localhost:8000"
echo ""
python3 main.py
