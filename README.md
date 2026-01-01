# Word文書テキスト変換ツール集

Word文書（.doc/.docx）をテキストファイル（.txt）に変換するための総合ツールセットです。単一ファイル変換から一括変換、GUI操作まで様々なニーズに対応しています。

---

## 📋 目次

- [概要](#概要)
- [推奨ツール（メイン使用）](#推奨ツールメイン使用)
- [全ファイル一覧と仕様](#全ファイル一覧と仕様)
- [セットアップ](#セットアップ)
- [使用方法](#使用方法)
- [推奨フロー](#推奨フロー)
- [トラブルシューティング](#トラブルシューティング)

---

## 📖 概要

このツール集は、Word文書からテキストを抽出するための複数のスクリプトを含んでいます。用途や環境に応じて最適なツールを選択できます。

### 主な機能
- ✅ .doc/.docx形式の両方に対応
- ✅ 一括変換（ディレクトリ単位）
- ✅ GUI操作可能
- ✅ 文字化け対策機能
- ✅ テキストクリーニング機能

---

## 🎯 推奨ツール（メイン使用）

### 🥇 **最推奨: `word_to_text_gui.py`** (GUI版・一括変換)

**用途**: **初心者向け・大量ファイルの一括変換**

**特徴**:
- ✅ GUI操作で簡単
- ✅ ディレクトリを指定して一括変換
- ✅ 進捗状況をリアルタイム表示
- ✅ ログで処理結果を確認可能
- ✅ サブディレクトリの再帰処理にも対応

**実行方法**:
```powershell
python word_to_text_gui.py
```

**使用シーン**:
- マニュアル集フォルダ全体を一度に変換したい
- 操作を視覚的に確認したい
- 変換結果をログで確認したい

---

### 🥈 **第2推奨: `word_to_text_converter.py`** (コマンドライン版・一括変換)

**用途**: **コマンドライン操作・自動化・一括変換**

**特徴**:
- ✅ コマンドラインで高速実行
- ✅ ディレクトリを指定して一括変換
- ✅ バッチ処理・自動化に最適
- ✅ 再帰処理のオン/オフ可能

**実行方法**:
```powershell
# 基本的な使い方
python word_to_text_converter.py "C:\path\to\マニュアル集"

# サブディレクトリを含めない場合
python word_to_text_converter.py "C:\path\to\マニュアル集" --no-recursive
```

**使用シーン**:
- GUIを使わずコマンドラインで処理したい
- バッチファイルやスクリプトから呼び出したい
- 定期的な自動処理に組み込みたい

---

### 🥉 **第3推奨: `simple_converter.py`** (単一ファイル変換)

**用途**: **個別ファイルの変換**

**特徴**:
- ✅ シンプルで理解しやすい
- ✅ .doc/.docx両対応
- ✅ エンコーディング指定可能

**実行方法**:
```powershell
# .docxファイルの変換
python simple_converter.py example.docx

# .docファイルの変換（エンコーディング指定）
python simple_converter.py example.doc --encoding=shift-jis
```

**使用シーン**:
- 1〜数個のファイルだけ変換したい
- ファイル形式を自動判別してほしい

---

## 📂 全ファイル一覧と仕様

### 🔵 メイン変換ツール（推奨）

| ファイル名 | 用途 | 対象 | 処理方法 | GUI | 推奨度 |
|-----------|------|------|---------|-----|--------|
| `word_to_text_gui.py` | 一括変換（GUI版） | ディレクトリ | 再帰処理 | ✅ | ⭐⭐⭐ |
| `word_to_text_converter.py` | 一括変換（CLI版） | ディレクトリ | 再帰処理 | ❌ | ⭐⭐⭐ |
| `simple_converter.py` | 単一ファイル変換 | 個別ファイル | .doc/.docx自動判別 | ❌ | ⭐⭐ |

---

### 🔸 専用変換ツール（個別用途）

| ファイル名 | 用途 | 対象形式 | 変換方法 | 推奨度 |
|-----------|------|---------|---------|--------|
| `docx_converter.py` | .docx専用変換 | .docx | docx2txt → python-docx | ⭐ |
| `doc_converter.py` | .doc専用変換 | .doc | Word COM | ⭐ |
| `word_converter.py` | 汎用変換 | .doc/.docx | docx2txt + バイナリ解析 | ⭐⭐ |
| `doc_to_txt.py` | 旧バージョン変換 | .doc/.docx | Word COM + docx2txt | ⭐ |
| `word-to-txt.py` | LibreOffice経由変換 | .doc/.docx | LibreOffice API | - |

#### 詳細仕様

##### `docx_converter.py` (.docx専用)
- **入力**: .docxファイル
- **出力**: .txtファイル（同名）
- **変換方法**: docx2txt → python-docx（フォールバック）
- **エンコーディング**: 指定可能（デフォルト: utf-8）
- **使用例**:
  ```powershell
  python docx_converter.py example.docx
  python docx_converter.py example.docx --encoding=shift-jis
  ```

##### `doc_converter.py` (.doc専用)
- **入力**: .docファイル
- **出力**: .txtファイル（同名）
- **変換方法**: Word COM（Windows専用）
- **エンコーディング**: 指定可能（デフォルト: utf-8）
- **前提条件**: Microsoft Wordインストール必須
- **使用例**:
  ```powershell
  python doc_converter.py example.doc
  python doc_converter.py example.doc --encoding=shift-jis
  ```

##### `word_converter.py` (汎用・バイナリ解析付き)
- **入力**: .doc/.docxファイル
- **出力**: .txtファイル（同名）
- **変換方法**: 
  1. docx2txt
  2. python-docx
  3. バイナリ直接解析（フォールバック）
- **特徴**: 変換失敗時にバイナリから直接テキスト抽出を試行
- **使用例**:
  ```powershell
  python word_converter.py example.doc
  python word_converter.py example.docx --encoding=utf-8
  ```

##### `doc_to_txt.py` (旧バージョン)
- **入力**: .doc/.docxファイル
- **出力**: .txtファイル（同名）
- **変換方法**: Word COM + docx2txt
- **特徴**: Windows環境専用、Word COMを優先使用
- **注意**: `word_to_text_converter.py`の旧版

##### `word-to-txt.py` (LibreOffice版)
- **入力**: .doc/.docxファイル
- **出力**: .txtファイル
- **変換方法**: LibreOffice API経由
- **前提条件**: LibreOfficeインストール必須
- **注意**: 設定が複雑なため非推奨

---

### 🔶 テキストクリーニングツール

| ファイル名 | 用途 | 処理内容 | 推奨度 |
|-----------|------|---------|--------|
| `cleanup_text.py` | 基本クリーニング | XMLタグ削除、制御文字削除、重複行削除 | ⭐⭐⭐ |
| `super_cleanup.py` | 高度クリーニング | 文字化け検出、段落整形 | ⭐⭐ |
| `final_cleanup.py` | 最終整形 | UTF-8正規化 | ⭐ |

#### 詳細仕様

##### `cleanup_text.py` （メイン推奨）
- **入力**: .txtファイル
- **出力**: `_cleaned.txt`ファイル
- **処理内容**:
  1. XMLタグの削除
  2. 制御文字・バイナリゴミの削除
  3. 日本語文字と英数字を含む行のみ抽出
  4. 重複行の削除
- **使用例**:
  ```powershell
  python cleanup_text.py example.txt
  python cleanup_text.py example.txt --encoding=utf-8 --output=result.txt
  ```

##### `super_cleanup.py`
- **入力**: .txtファイル
- **出力**: 指定したファイル名
- **処理内容**:
  - 文字化けパターンの自動検出
  - 段落単位での整形
  - 日本語らしさチェック
- **使用例**:
  ```powershell
  # スクリプト内でファイル名を指定
  python super_cleanup.py
  ```

##### その他クリーニングツール
- `final_cleanup.py`: UTF-8正規化処理
- `final_cleaner.py`: 最終整形用
- `convert_utf8.py`: エンコーディング変換
- `direct_utf8_fix.py`: UTF-8問題修正
- `enhanced_utf8_fix.py`: 拡張UTF-8修正
- `fix_txt.py`: テキスト修正

**注意**: これらは特定の問題対応用で、通常は`cleanup_text.py`で十分です。

---

## 🔧 セットアップ

### 前提条件

- **Python 3.6以上**
- **Microsoft Word** （.doc変換時に必要）
- **Windows OS** （推奨、Linuxでも一部動作可能）

### インストール

1. **リポジトリのダウンロード**
   ```powershell
   cd "C:\Users\looka\OneDrive\ドキュメント\job\wordをtxtに変換"
   ```

2. **必要なライブラリのインストール**
   ```powershell
   pip install -r requirements.txt
   ```

### requirements.txt の内容
```
python-docx>=0.8.11
pywin32>=305 
pypandoc>=1.11
comtypes>=1.1.14
```

---

## 🚀 使用方法

### パターン1: GUI版で一括変換（最も簡単）

```powershell
# GUIツールを起動
python word_to_text_gui.py
```

**操作手順**:
1. 「ディレクトリを選択」ボタンをクリック
2. マニュアル集フォルダを選択
3. 「サブディレクトリも含める」にチェック（必要に応じて）
4. 「変換開始」ボタンをクリック
5. 進捗を確認しながら待機
6. 完了後、ログで結果を確認

---

### パターン2: コマンドラインで一括変換

```powershell
# ディレクトリ全体を再帰的に変換
python word_to_text_converter.py "C:\path\to\マニュアル集"

# サブディレクトリを含めない場合
python word_to_text_converter.py "C:\path\to\マニュアル集" --no-recursive
```

---

### パターン3: 単一ファイルの変換

```powershell
# 自動判別で変換
python simple_converter.py example.doc

# .docxファイル
python simple_converter.py example.docx

# エンコーディング指定
python simple_converter.py example.doc --encoding=shift-jis
```

---

### パターン4: 変換後のテキストクリーニング

```powershell
# 変換後のテキストをクリーニング
python cleanup_text.py example.txt

# 出力先を指定
python cleanup_text.py example.txt --output=cleaned_example.txt

# エンコーディング指定
python cleanup_text.py example.txt --encoding=shift-jis
```

---

## 🎓 推奨フロー

### 基本フロー（ほとんどの場合に適用）

```
ステップ1: GUI版で一括変換
  ↓ word_to_text_gui.py
  
ステップ2: （必要に応じて）クリーニング
  ↓ cleanup_text.py
  
完了: .txtファイル取得
```

### 詳細フロー（問題がある場合）

```
1. GUI版で一括変換
   python word_to_text_gui.py
   ↓
   
2. 変換結果を確認
   ↓
   
3-A. 文字化けがある場合
   python cleanup_text.py [ファイル名.txt]
   ↓
   
3-B. それでも文字化けが残る場合
   python super_cleanup.py
   ↓
   
3-C. エンコーディングの問題の場合
   python simple_converter.py [ファイル名] --encoding=shift-jis
   ↓
   
4. 完了
```

### ケース別推奨ツール

| ケース | 推奨ツール | 理由 |
|--------|----------|------|
| 初めて使う | `word_to_text_gui.py` | GUI操作で直感的 |
| 大量ファイル | `word_to_text_gui.py` | 進捗確認が容易 |
| 自動化・バッチ処理 | `word_to_text_converter.py` | コマンドライン対応 |
| 1〜数ファイル | `simple_converter.py` | シンプルで高速 |
| 文字化けがある | `cleanup_text.py` | クリーニング機能充実 |
| .doc形式のみ | `doc_converter.py` | Word COM使用で高品質 |
| .docx形式のみ | `docx_converter.py` | 軽量で高速 |

---

## ⚠️ 注意事項

### 共通の注意事項

1. **出力ファイルの上書き**
   - 同名の.txtファイルが存在する場合、確認なしで上書きされます
   - 重要なファイルはバックアップを取ってください

2. **書式情報の喪失**
   - テキスト変換では、フォント、色、レイアウトなどの書式情報は失われます
   - 表や画像は変換されない場合があります

3. **エンコーディング**
   - デフォルトはUTF-8
   - 古いファイルでは`shift-jis`や`cp932`が適している場合があります

### .doc変換の注意事項

- **Microsoft Word必須**: .doc変換にはWordがインストールされている必要があります
- **COM使用**: Windows環境でのみ動作します
- **処理時間**: .docx変換より時間がかかる場合があります

### GUI版の注意事項

- **メモリ使用量**: 大量のファイル処理時はメモリを多く使用します
- **処理中断**: 「キャンセル」ボタンで中断可能ですが、処理中のファイルは完了まで待機されます

---

## 🐛 トラブルシューティング

### エラー: `ImportError: No module named 'win32com'`

**原因**: pywin32がインストールされていない

**解決方法**:
```powershell
pip install pywin32
```

---

### エラー: `FileNotFoundError`

**原因**: 指定したファイルまたはディレクトリが存在しない

**解決方法**:
- ファイルパスを確認
- 絶対パスで指定
- パスに`r`プレフィックスを付ける（例: `r"C:\Users\..."`）

---

### 問題: 文字化けが発生する

**原因**: エンコーディングの不一致

**解決方法（優先順）**:
1. `cleanup_text.py`でクリーニング
   ```powershell
   python cleanup_text.py example.txt
   ```

2. エンコーディングを変更して再変換
   ```powershell
   python simple_converter.py example.doc --encoding=shift-jis
   ```

3. `super_cleanup.py`で高度クリーニング
   ```powershell
   python super_cleanup.py
   ```

---

### 問題: .doc変換に失敗する

**原因**: Word COMの初期化失敗

**解決方法**:
1. Wordを一度起動して正常に動作するか確認
2. Wordを終了してから再度実行
3. `word_converter.py`でバイナリ解析を試す
   ```powershell
   python word_converter.py example.doc
   ```

---

### 問題: GUI版が起動しない

**原因**: tkinterライブラリの問題

**解決方法**:
- Python再インストール時に「tcl/tk」を含める
- コマンドライン版を使用
  ```powershell
  python word_to_text_converter.py "C:\path\to\dir"
  ```

---

### 問題: 処理が遅い

**原因**: 大量のファイル処理、またはCOM使用

**対策**:
- サブディレクトリを分割して処理
- .doc→.docx変換を事前に実施（Wordで一括変換）
- `simple_converter.py`で.docxのみ処理

---

### 問題: 特定のファイルだけ変換失敗する

**原因**: ファイル破損、または特殊な形式

**解決方法**:
1. Wordで直接開けるか確認
2. Wordで「名前を付けて保存」で再保存
3. 別のツールで変換を試す
   ```powershell
   python word_converter.py example.doc
   ```

---

## 📚 使用ライブラリ

| ライブラリ | 用途 | 必須/任意 |
|-----------|------|----------|
| `python-docx` | .docxファイル読み込み | 必須 |
| `docx2txt` | .docxテキスト抽出 | 必須 |
| `pywin32` | Word COM操作 | .doc変換時必須 |
| `pypandoc` | Pandoc経由変換 | 任意 |
| `comtypes` | LibreOffice COM | 任意 |
| `tkinter` | GUI表示 | GUI版で必須 |

---

## 🔄 ファイル間の関係図

```
【ユーザー】
    ↓
┌────────────────────────────┐
│  GUI版 or CLI版で一括変換   │
│  - word_to_text_gui.py      │
│  - word_to_text_converter.py│
└───────────┬────────────────┘
            │
            ↓
┌────────────────────────────┐
│  内部で各変換ツールを使用    │
│  - docx_converter.py        │
│  - doc_converter.py         │
│  - word_converter.py        │
└───────────┬────────────────┘
            │
            ↓
      【.txtファイル生成】
            │
            ↓（文字化けがある場合）
┌────────────────────────────┐
│  クリーニングツール          │
│  - cleanup_text.py          │
│  - super_cleanup.py         │
└───────────┬────────────────┘
            │
            ↓
    【最終的な.txtファイル】
```

---

## 📞 サポート

不明点や改善要望があれば、開発担当者までご連絡ください。

---

## 📜 ライセンス

社内ツールとして使用してください。

---

**最終更新**: 2026年1月1日 