@echo off
setlocal enabledelayedexpansion

:: ===============================
:: 【1】日付を取得（当日分→前日分に変更）
:: ===============================
for /f %%I in ('powershell -command "(Get-Date).AddDays(-1).ToString('yyyyMMdd')"') do set YYYYMMDD=%%I

:: ===============================
:: 【2】各種パスを定義
:: ===============================
set BASE_DIR=C:\Users\rep03\Desktop\Pythonファイル\アップデートとビズコラボ_進捗Report
set PYTHON_EXEC=C:\Users\rep03\Desktop\Pythonファイル\Python実行環境\python.exe
set LOG_DIR=%BASE_DIR%\log
set LOG_FILE=%LOG_DIR%\アップデートとビズコラボ進捗レポート実行_%YYYYMMDD%分.log

:: ===============================
:: 【3】ログフォルダ作成
:: ===============================
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: ===============================
:: 【4】ログヘッダ出力
:: ===============================
echo ==================================================== >> "%LOG_FILE%"
echo [%date% %time%] [INFO] アップデートとビズコラボ_日次処理 開始 >> "%LOG_FILE%"
echo ---------------------------------------------------- >> "%LOG_FILE%"
echo [%date% %time%] [INFO] 実行ディレクトリ: %BASE_DIR% >> "%LOG_FILE%"
echo [%date% %time%] [INFO] 使用するPython: %PYTHON_EXEC% >> "%LOG_FILE%"
echo ==================================================== >> "%LOG_FILE%"

:: ===============================
:: 【5】Pythonの存在確認
:: ===============================
"%PYTHON_EXEC%" --version >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] Pythonが見つかりません。環境設定を確認してください。 >> "%LOG_FILE%"
    exit /b 1
)

:: ===============================
:: 【6】STEP 01
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 01] 日時レポート計測値取得と各ID進捗表に反映 >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\01_日時レポート計測値取得と各ID進捗表に反映.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 01 実行失敗 >> "%LOG_FILE%"
)

:: ===============================
:: 【7】STEP 02
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 02] アップデートとビズコラボ_fam8進捗Report取得(AX-AD) >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\02_アップデートとビズコラボ_fam8進捗Report取得(AX-AD).py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 02 実行失敗 >> "%LOG_FILE%"
)

:: ===============================
:: 【8】STEP 03
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 03] CSVを集計用シート（AX-AD）へアップロード >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\03_CSVを集計用シート（AX-AD）へアップロード.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 03 実行失敗 >> "%LOG_FILE%"
)

:: ===============================
:: 【9】STEP 04
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 04] 今月シートのAX-AD表に関数挿入と値のみ変換 >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\04_今月シートのAX-AD表に関数挿入と値のみ変換.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 04 実行失敗 >> "%LOG_FILE%"
)

:: ===============================
:: 【10】STEP 05
:: ===============================
echo --------------------------------------- >> "%LOG_FILE%"
echo [STEP 05] マイム表_金額計算結果のみ記入 >> "%LOG_FILE%"
echo --------------------------------------- >> "%LOG_FILE%"
"%PYTHON_EXEC%" "%BASE_DIR%\05_マイム表_金額計算結果のみ記入.py" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] STEP 05 実行失敗 >> "%LOG_FILE%"
)

:: ===============================
:: 【11】処理完了
:: ===============================
echo ==================================================== >> "%LOG_FILE%"
echo [%date% %time%] [INFO] 全処理完了 >> "%LOG_FILE%"
echo ==================================================== >> "%LOG_FILE%"

exit /b 0
