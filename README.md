# Excel Differ Git

Git統合されたExcel差分表示ツール。Excelファイルの変更を行単位で比較し、挿入・削除・更新を正確に検出します。

## 特徴

- **Git統合**: Gitコミット間またはワーキングツリーとの差分を表示
- **行単位比較**: 行の挿入・削除があっても、内容ベースで正確に差分を検出
- **複数シート対応**: ブック内のすべてのシートを比較
- **複数出力形式**: テキスト（標準）、CSV形式に対応
- **柔軟な比較**: Git統合モードと直接ファイル比較モードをサポート

## インストール

```bash
# リポジトリをクローン
git clone https://github.com/k-tanaka-522/EXCEL-Differ-Git.git
cd EXCEL-Differ-Git

# 依存関係をインストール
pip install -r requirements.txt

# 開発モードでインストール（推奨）
pip install -e .
```

## 使い方

### Git統合モード

```bash
# 最新コミットと1つ前のコミットを比較（デフォルト）
excel-diff myfile.xlsx

# 特定のコミット間を比較
excel-diff --from HEAD~2 --to HEAD myfile.xlsx
excel-diff --from abc123 --to def456 myfile.xlsx

# ワーキングツリー（未コミットの変更）との比較
excel-diff --working-tree myfile.xlsx
excel-diff --from HEAD~1 --working-tree myfile.xlsx
```

### ファイル直接比較モード

```bash
# 2つのExcelファイルを直接比較
excel-diff --old old_version.xlsx --new new_version.xlsx old_version.xlsx
```

### 出力形式

```bash
# テキスト形式（デフォルト、標準出力）
excel-diff myfile.xlsx

# CSV形式で出力
excel-diff --format csv myfile.xlsx

# ファイルに保存
excel-diff --format csv --output diff.csv myfile.xlsx
excel-diff --format text -o diff.txt myfile.xlsx
```

## 出力例

### テキスト形式

```
======================================================================
Excel Diff: config.xlsx
Comparing: config.xlsx (a1b2c3d) ⟷ config.xlsx (e4f5g6h)
======================================================================

[Sheet: 基本設定]
  Row 3 MODIFIED
    Column B: "8080" → "8081"

[Sheet: 監視設定]
  Row 5 DELETED
    - 監視ID_003 | ディスク監視 | 95

  Row 7 ADDED
    + 監視ID_004 | ネットワーク監視 | 100

  Row 8 MODIFIED
    Column C: "90" → "85"

======================================================================
Summary
======================================================================
Sheets modified: 2
Rows added: 1
Rows deleted: 1
Rows modified: 2
```

### CSV形式

```csv
type,sheet,old_row,new_row,column,cell,old_value,new_value,description
cell_modified,基本設定,3,3,B,B3,8080,8081,Cell modified
row_deleted,監視設定,5,,,,監視ID_003|ディスク監視|95,,Row deleted
row_added,監視設定,,7,,,,"監視ID_004|ネットワーク監視|100",Row added
cell_modified,監視設定,8,8,C,C8,90,85,Cell modified
```

## 差分検出アルゴリズム

このツールは内容ベースの行マッチングを使用しています：

1. **完全一致**: 行の内容が完全に一致する行を最初に検出
2. **類似行検出**: 残った行から50%以上のセルが一致する行を「変更」として検出
3. **追加・削除**: マッチしなかった行を「追加」または「削除」として分類

これにより、複数行の挿入・削除があっても、実際の変更のみを正確に抽出できます。

## プロジェクト構造

```
ExcelDifferGit/
├── excel_differ/
│   ├── __init__.py
│   ├── cli.py           # CLIエントリーポイント
│   ├── git_handler.py   # Git操作
│   ├── excel_reader.py  # Excel読み込み
│   ├── differ.py        # 差分検出アルゴリズム
│   └── formatter.py     # 出力フォーマット
├── tests/               # テスト（今後追加予定）
├── pyproject.toml       # プロジェクト設定
├── requirements.txt     # 依存関係
└── README.md
```

## 依存ライブラリ

- **openpyxl** (>=3.1.0): Excel読み書き
- **GitPython** (>=3.1.0): Git操作
- **click** (>=8.1.0): CLI構築

## 対応フォーマット

- `.xlsx` (Excel 2007以降)
- `.xlsm` (マクロ有効Excel)

## ライセンス

MIT License

## 今後の機能追加予定

- [ ] JSON出力形式
- [ ] 書式変更の検出
- [ ] 数式変更の検出
- [ ] 特定シートのみ比較するフィルタ機能
- [ ] 空行を無視するオプション
- [ ] キー列を指定した高度なマッチング
