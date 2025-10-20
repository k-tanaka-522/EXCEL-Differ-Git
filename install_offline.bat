@echo off
REM Excel Differ Git - 完全オフラインインストールスクリプト
REM このスクリプトはインターネット接続なしで実行できます

echo ====================================================================
echo Excel Differ Git - オフラインインストール
echo ====================================================================
echo.

REM 依存関係をインストール
echo [1/2] 依存ライブラリをインストール中...
py -m pip install --no-index --find-links=vendor -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo エラー: 依存ライブラリのインストールに失敗しました
    pause
    exit /b 1
)
echo.

REM アプリ本体をインストール
echo [2/2] Excel Differ本体をインストール中...
py -m pip install --no-index --find-links=vendor -e .
if %ERRORLEVEL% NEQ 0 (
    echo エラー: アプリ本体のインストールに失敗しました
    pause
    exit /b 1
)
echo.

echo ====================================================================
echo インストール完了！
echo ====================================================================
echo.
echo 使い方:
echo   py -m excel_differ.cli --old old.xlsx --new new.xlsx old.xlsx
echo.
echo または、Scriptsディレクトリにexcel-diff.exeが作成されました
echo （PATHに追加すればexcel-diffコマンドで実行可能）
echo.
pause
