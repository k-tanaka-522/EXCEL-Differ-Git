"""
依存ライブラリをダウンロードするスクリプト

このスクリプトは、pip install が使えない環境のために
必要なライブラリとその依存関係をすべてダウンロードします。

使い方:
    python download_dependencies.py
    または
    py download_dependencies.py
    または
    python3 download_dependencies.py
"""
import subprocess
import sys
import os
import shutil

def find_python_command():
    """利用可能なPythonコマンドを検出"""
    commands = ['py', 'python', 'python3']
    for cmd in commands:
        if shutil.which(cmd):
            return cmd
    return None

def download_dependencies():
    """requirements.txt の依存関係をすべてダウンロード"""

    # vendorディレクトリを作成
    vendor_dir = os.path.join(os.path.dirname(__file__), 'vendor')
    os.makedirs(vendor_dir, exist_ok=True)

    print(f"依存ライブラリを {vendor_dir} にダウンロードしています...")
    print("=" * 60)

    try:
        # pip download コマンドで依存関係をすべてダウンロード
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'download',
            '-r', 'requirements.txt',
            '-d', vendor_dir,
            '--no-cache-dir'
        ])

        print("=" * 60)
        print(f"✓ ダウンロード完了: {vendor_dir}")
        print()
        print("次のステップ:")
        print("1. このリポジトリ全体をZIPでパッケージ化")
        print("2. オフライン環境で解凍")
        print("3. install_offline.bat (Windows) または install_offline.sh (Linux/Mac) を実行")

    except subprocess.CalledProcessError as e:
        print(f"エラー: ダウンロードに失敗しました: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    download_dependencies()
