@echo off
cd /d "%~dp0\..\Python実行環境"

echo [INFO] pip導入処理を開始...

:: pip 導入
python.exe get-pip.py
if %errorlevel% neq 0 (
    echo [ERROR] pipの導入に失敗しました
    pause
    exit /b 1
)

:: ライブラリ一括インストール
python.exe -m pip install --upgrade pip
python.exe -m pip install -r "%~dp0\requirements.txt"
if %errorlevel% neq 0 (
    echo [ERROR] ライブラリのインストールに失敗しました
    pause
    exit /b 1
)

echo [SUCCESS] pipとライブラリのセットアップが完了しました
pause
exit /b 0
