@echo off
REM Windows用オフラインインストールスクリプト
REM このスクリプトは、vendorディレクトリ内のwhlファイルから依存関係をインストールします

echo ========================================
echo Excel Differ Git - オフラインインストール
echo ========================================
echo.

REM vendorディレクトリの存在確認
if not exist "vendor" (
    echo エラー: vendorディレクトリが見つかりません
    echo download_dependencies.py を実行してください
    exit /b 1
)

REM Pythonコマンドを検出
set PYTHON_CMD=
where py >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :python_found
)
where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :python_found
)
where python3 >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3
    goto :python_found
)

echo エラー: Pythonが見つかりません
echo Python 3.8以上をインストールしてください
exit /b 1

:python_found
echo Pythonコマンド: %PYTHON_CMD%
echo.

echo vendorディレクトリから依存ライブラリをインストール中...
echo.

REM vendorディレクトリから依存関係をインストール
%PYTHON_CMD% -m pip install --no-index --find-links=vendor -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo エラー: インストールに失敗しました
    exit /b 1
)

echo.
echo 本体をインストール中...
%PYTHON_CMD% -m pip install --no-index --find-links=vendor -e .

if %errorlevel% neq 0 (
    echo.
    echo エラー: 本体のインストールに失敗しました
    exit /b 1
)

echo.
echo ========================================
echo インストール完了！
echo ========================================
echo.
echo 使い方:
echo   excel-diff myfile.xlsx
echo.
echo 詳細は README.md を参照してください
echo.

pause
