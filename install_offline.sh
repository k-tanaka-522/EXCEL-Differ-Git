#!/bin/bash
# Linux/Mac用オフラインインストールスクリプト
# このスクリプトは、vendorディレクトリ内のwhlファイルから依存関係をインストールします

echo "========================================"
echo "Excel Differ Git - オフラインインストール"
echo "========================================"
echo

# vendorディレクトリの存在確認
if [ ! -d "vendor" ]; then
    echo "エラー: vendorディレクトリが見つかりません"
    echo "download_dependencies.py を実行してください"
    exit 1
fi

# Pythonコマンドを検出
PYTHON_CMD=""
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
elif command -v py &> /dev/null; then
    PYTHON_CMD="py"
else
    echo "エラー: Pythonが見つかりません"
    echo "Python 3.8以上をインストールしてください"
    exit 1
fi

echo "Pythonコマンド: $PYTHON_CMD"
echo

echo "vendorディレクトリから依存ライブラリをインストール中..."
echo

# vendorディレクトリから依存関係をインストール
$PYTHON_CMD -m pip install --no-index --find-links=vendor -r requirements.txt

if [ $? -ne 0 ]; then
    echo
    echo "エラー: インストールに失敗しました"
    exit 1
fi

echo
echo "本体をインストール中..."
$PYTHON_CMD -m pip install --no-index --find-links=vendor -e .

if [ $? -ne 0 ]; then
    echo
    echo "エラー: 本体のインストールに失敗しました"
    exit 1
fi

echo
echo "========================================"
echo "インストール完了！"
echo "========================================"
echo
echo "使い方:"
echo "  excel-diff myfile.xlsx"
echo
echo "詳細は README.md を参照してください"
echo
