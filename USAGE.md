# Excel Differ Git - 使い方ガイド

## 基本的な使い方

### 1. ローカルファイル2つの直接比較（Git不要）

```bash
# 2つのExcelファイルを直接比較
py -m excel_differ.cli --old old_version.xlsx --new new_version.xlsx old_version.xlsx

# テキスト形式で標準出力
py -m excel_differ.cli --old test_config_v1.xlsx --new test_config_v2.xlsx test_config_v1.xlsx

# CSV形式でファイルに保存
py -m excel_differ.cli --old old.xlsx --new new.xlsx --format csv --output diff.csv old.xlsx
```

**注意**: `--old` と `--new` を使う場合、最後の引数（FILE）はどちらかのファイルパスを指定してください。

### 2. Git統合モード

```bash
# 最新コミットと1つ前のコミットを比較
py -m excel_differ.cli myfile.xlsx

# 特定のコミット間を比較
py -m excel_differ.cli --from HEAD~2 --to HEAD myfile.xlsx
py -m excel_differ.cli --from abc123 --to def456 myfile.xlsx

# ワーキングツリー（未コミットの変更）との比較
py -m excel_differ.cli --working-tree myfile.xlsx
py -m excel_differ.cli --from HEAD~1 --working-tree myfile.xlsx
```

### 3. 出力形式

#### テキスト形式（デフォルト）
```bash
py -m excel_differ.cli --old v1.xlsx --new v2.xlsx v1.xlsx
```

出力例:
```
======================================================================
Excel Diff: test_config_v2.xlsx
Comparing: test_config_v1.xlsx <-> test_config_v2.xlsx
======================================================================

[Sheet: 基本設定]
  Row 3 MODIFIED
    Column B: "8080" -> "8081"

[Sheet: 監視設定]
  Row 5 DELETED
    - MON_003|ディスク監視|95

  Row 7 ADDED
    + MON_004|ネットワーク監視|100
```

#### CSV形式
```bash
py -m excel_differ.cli --old v1.xlsx --new v2.xlsx --format csv v1.xlsx
```

出力例:
```csv
type,sheet,old_row,new_row,column,cell,old_value,new_value,description
cell_modified,基本設定,3,3,B,B3,8080,8081,Cell modified
row_deleted,監視設定,5,,,,MON_003|ディスク監視|95,,Row deleted
row_added,監視設定,,7,,,,MON_004|ネットワーク監視|100,Row added
```

## よくある使用例

### HinemosUtilityの設定ファイル比較

```bash
# 編集前後の設定ファイルを比較
py -m excel_differ.cli --old hinemos_before.xlsx --new hinemos_after.xlsx --format csv --output changes.csv hinemos_before.xlsx

# 変更サマリを確認
py -m excel_differ.cli --old hinemos_before.xlsx --new hinemos_after.xlsx hinemos_before.xlsx | grep Summary -A 10
```

### 定期バックアップの差分確認

```bash
# 週次バックアップ間の差分
py -m excel_differ.cli --old backup_week1.xlsx --new backup_week2.xlsx --format csv backup_week1.xlsx > weekly_changes.csv
```

### レビュー用の差分レポート作成

```bash
# テキスト形式でレポート作成
py -m excel_differ.cli --old original.xlsx --new reviewed.xlsx --output review_report.txt original.xlsx

# CSV形式でレポート作成（Excelで開ける）
py -m excel_differ.cli --old original.xlsx --new reviewed.xlsx --format csv --output review_report.csv original.xlsx
```

## オプション一覧

| オプション | 短縮形 | 説明 |
|-----------|--------|------|
| `--old` | - | 比較元のExcelファイル（直接比較モード） |
| `--new` | - | 比較先のExcelファイル（直接比較モード） |
| `--from` | - | 比較元のGitコミット（Git統合モード） |
| `--to` | - | 比較先のGitコミット（デフォルト: HEAD） |
| `--working-tree` | - | ワーキングツリーと比較 |
| `--format` | - | 出力形式（text, csv） |
| `--output` | `-o` | 出力ファイルパス |

## 差分の見方

### 変更タイプ

- **ADDED (追加)**: 新しく追加された行
- **DELETED (削除)**: 削除された行
- **MODIFIED (変更)**: 内容が変更された行
- **SHEET_ADDED (シート追加)**: 新しく追加されたシート
- **SHEET_DELETED (シート削除)**: 削除されたシート

### 行の対応付けアルゴリズム

1. **完全一致**: 行の内容が完全に一致する行をまず検出
2. **類似度判定**: 50%以上のセルが一致する行を「変更」とみなす
3. **追加・削除**: マッチしなかった行を「追加」または「削除」に分類

このアルゴリズムにより、複数行の挿入・削除があっても、実際の変更のみを正確に抽出できます。

## トラブルシューティング

### 文字化けする場合

コンソール出力では日本語が文字化けすることがあります。CSV形式で出力してExcelで開いてください。

```bash
py -m excel_differ.cli --old v1.xlsx --new v2.xlsx --format csv --output diff.csv v1.xlsx
```

### ファイルが見つからない

絶対パスまたは相対パスを正しく指定してください。

```bash
# 相対パス
py -m excel_differ.cli --old ./data/old.xlsx --new ./data/new.xlsx ./data/old.xlsx

# 絶対パス
py -m excel_differ.cli --old "C:\data\old.xlsx" --new "C:\data\new.xlsx" "C:\data\old.xlsx"
```

### 大きなファイルの処理が遅い

1000行×1000列のファイルでも処理可能ですが、時間がかかる場合があります。
処理中は気長にお待ちください。
