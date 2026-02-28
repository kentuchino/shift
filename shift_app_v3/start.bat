@echo off
chcp 65001 > nul
echo ============================================================
echo  シフト表自動作成アプリ v3.0 - 起動中...
echo ============================================================
echo.

:: Pythonのチェック
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo.
    echo 以下のどちらかの方法で起動してください:
    echo.
    echo 【方法1: Renderクラウド版を使う】
    echo   ブラウザで以下のURLを開いてください:
    echo   https://shift-app-xxxxx.onrender.com
    echo   （URLはデプロイ手順書.mdを参照）
    echo.
    echo 【方法2: Python環境を用意する】
    echo   https://www.python.org/ からインストール
    echo.
    pause
    exit /b 1
)

:: ライブラリのインストール確認
pip show fastapi > nul 2>&1
if %errorlevel% neq 0 (
    echo ライブラリをインストールしています（初回のみ）...
    pip install -r requirements.txt --quiet
)

echo.
echo ブラウザが自動で開きます。開かない場合は以下にアクセス:
echo http://localhost:8000
echo.
echo 終了するにはこのウィンドウを閉じてください。
echo.

python main.py

pause
