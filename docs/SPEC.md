# 名簿帳票ツール 開発仕様書

> **バージョン**: 3.0（統合版）  
> **対象**: Claude Code + VSCode によるコーディングエージェント実装  
> **構成方針**: 各章が「仕様 → 実装コード → 落とし穴」をワンセットで含む

---

## 第1章 プロジェクト概要

### 1.1 背景と目的

那覇市立小中学校では、校務支援システム（C4th / EDUCOM）からエクスポートした Excel 名簿データを基に、名札・名列表・台帳・個票などの帳票を手作業で作成している。本プロジェクトでは、名簿データの読み込みからテンプレートへの差し込み、Excel ファイル出力までを自動化する Windows デスクトップアプリケーションを開発する。

### 1.2 前提条件と制約

- **IPAmj明朝フォント要件**: 児童生徒の氏名に外字（異体字）が含まれるケースがあり、58,712 文字を収録する IPAmj明朝で正確に表示・印刷する。C4th の「正式名前」列に IVS 付き文字が含まれる。本アプリは Excel ファイルにフォント名メタデータを書き込み、学校 PC にインストール済みの IPAmj明朝で描画・印刷する方式。
- **動作環境**: Windows 10/11、Microsoft Excel 2016 以降、Python ランタイム不要（exe 化配布）。
- **ネットワーク**: 初回起動・自動更新時のみ Google ドライブへの HTTPS 接続。帳票生成自体はオフライン動作。

### 1.3 技術スタック

| 項目 | 技術 | バージョン | 用途 |
|------|------|-----------|------|
| 言語 | Python | 3.11+ | メインロジック |
| GUI | CustomTkinter | 5.x | モダンな UI 構築 |
| Excel 操作 | openpyxl | 3.1+ | テンプレート読み込み・データ書き込み・xlsx 出力 |
| Excel 操作（補助） | xlwings（オプション） | 0.30+ | Excel COM 経由の直接印刷 |
| データ処理 | pandas | 2.x | 名簿データ操作 |
| exe 化 | PyInstaller | 6.x | パッケージング |
| HTTP 通信 | requests | 2.x | Google Drive 自動更新 |
| 画像 | Pillow | 9.0+ | openpyxl の画像操作に必須 |

### 1.4 ソースコード構成

```
meibo_tool/
  ├── main.py                  # エントリーポイント
  ├── gui/
  │   ├── app.py               # メインウィンドウ（CustomTkinter）
  │   ├── frames/
  │   │   ├── import_frame.py  # 名簿読込パネル
  │   │   ├── select_frame.py  # 学年・組・テンプレート選択パネル
  │   │   ├── preview_frame.py # データプレビュー＋手動カラム追加
  │   │   └── output_frame.py  # 生成・印刷パネル
  │   └── dialogs/
  │       ├── mapping_dialog.py    # カラムマッピング画面
  │       ├── add_column_dialog.py # 手動カラム追加ダイアログ
  │       └── update_dialog.py     # 更新確認ダイアログ
  ├── core/
  │   ├── importer.py          # C4th Excel データ読込
  │   ├── mapper.py            # カラム名マッピング（確定済み辞書）
  │   ├── generator.py         # テンプレートへのデータ差込処理
  │   ├── updater.py           # Google Drive 自動更新
  │   ├── printer.py           # 直接印刷（オプション）
  │   └── config.py            # 設定ファイル管理
  ├── templates/
  │   ├── template_registry.py # テンプレートメタデータ
  │   └── generators/          # テンプレート生成スクリプト
  │       ├── generate_all.py
  │       ├── gen_nafuda_plain.py
  │       ├── gen_nafuda_decorated.py
  │       ├── gen_nafuda_1nen.py
  │       ├── gen_meireihyo.py
  │       ├── gen_shirabehyo.py
  │       ├── gen_shuuryo_daicho.py
  │       ├── gen_sotsugyou_daicho.py
  │       ├── gen_katei_chousahyo.py
  │       └── gen_gakkyuu_kojihyo.py
  ├── utils/
  │   ├── wareki.py            # 西暦→和暦変換
  │   ├── address.py           # 住所結合ヘルパー
  │   └── font_helper.py       # IPAmj明朝フォント設定
  └── tests/
      ├── test_wareki.py
      ├── test_address.py
      ├── test_mapper.py
      ├── test_importer.py
      ├── test_generator.py
      ├── test_updater.py
      ├── fixtures/
      │   ├── dummy_c4th.xlsx
      │   └── expected/
      └── generate_dummy.py
```

### 1.5 配布フォルダ構成（--onedir 方式）

```
名簿帳票ツール/                    ← フォルダ全体を共有
  ├── 名簿帳票ツール.exe            # アプリ本体
  ├── _internal/                    # PyInstaller 内部ファイル（触らない）
  │   ├── customtkinter/
  │   └── ... (DLL群)
  ├── config.json                   # 設定ファイル
  ├── テンプレート/                  # Excel テンプレート群
  │   ├── 名札_通常.xlsx
  │   ├── 名札_装飾あり.xlsx
  │   ├── 名札_1年生用.xlsx
  │   ├── 掲示用名列表.xlsx
  │   ├── 調べ表.xlsx
  │   ├── 修了台帳.xlsx
  │   ├── 卒業台帳.xlsx
  │   ├── 家庭調査票.xlsx
  │   ├── 学級編成用個票.xlsx
  │   └── assets/
  │       └── flower_border.png
  └── 出力/                         # 生成ファイル出力先
```

### 1.6 処理フロー全体像

```
① exe ダブルクリック → アプリ起動
② [バックグラウンド] updater.py が Google Drive の version.json をチェック
③ 更新がある場合 → ダイアログで通知 → 同意でダウンロード＆更新
④ 「名簿読込」ボタンで C4th エクスポート Excel を選択
⑤ importer.py がヘッダー自動検出 → mapper.py でカラムマッピング
⑥ 【必須ステップ】必須情報入力パネルが表示される：
    - 学年：C4th から自動セット（編集可）
    - 組：教員が必ず選択（未設定だと先に進めない）
    - 出席番号：「自動連番」ボタンで一括付与（個別編集可）
⑦ 「確定して進む」ボタン → テンプレート選択が有効化
⑧ [任意]「列を追加」で担任名・証書番号等を追加
⑨ テンプレート一覧から帳票を選択
⑩ 「Excel 生成」ボタン → generator.py が xlsx 出力
⑪ 出力フォルダが開き、教員が Excel で印刷
```

### 1.7 config.json 仕様

```json
{
  "app_version": "1.0.0",
  "school_name": "那覇市立天久小学校",
  "school_type": "elementary",
  "template_dir": "./テンプレート",
  "output_dir": "./出力",
  "default_font": "IPAmj明朝",
  "fiscal_year": 2025,
  "graduation_cert_start_number": 1,
  "homeroom_teachers": {
    "1-1": "山田先生",
    "1-2": "鈴木先生"
  },
  "update": {
    "version_file_id": "1AbCd_version_json_gdrive_file_id",
    "check_on_startup": true,
    "current_app_version": "1.0.0",
    "current_template_version": "1.0.0"
  },
  "column_mappings": {
    "last_used": {}
  },
  "manual_columns": {
    "last_file_hash": "abc123",
    "mandatory": {
      "組": {"value": "1"},
      "出席番号": {"type": "sequential", "start": 1}
    },
    "optional": {
      "担任名": {"type": "uniform", "value": "山田先生"}
    }
  },
  "recent_files": []
}
```

### 1.8 requirements.txt

```
customtkinter>=5.2.0
openpyxl>=3.1.2
pandas>=2.0.0
requests>=2.28.0
Pillow>=9.0.0
pyinstaller>=6.0.0    # ビルド時のみ
```

### 1.9 動作環境要件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11（64bit） |
| Excel | Microsoft Excel 2016 以降 |
| フォント | IPAmj明朝インストール済み |
| ネットワーク | HTTPS 接続可能（オフラインでも帳票生成は可能） |
| ディスク | exe: 50〜100MB、テンプレート: 5MB |
| Python | 不要 |

---

## 第2章 データモデル

### 2.1 C4th エクスポート Excel — 確定済みヘッダー（全50カラム）

以下は C4th（EDUCOM）からエクスポートされる実際のヘッダー（2026年2月確認済み）。

| # | C4th ヘッダー（実際の列名） | 内部論理名 | 用途・備考 |
|---|--------------------------|-----------|-----------|
| 1 | 生徒コード | 生徒コード | C4th 内部 ID。データ突合に使用 |
| 2 | 学年 | 学年 | 1〜6（小）、1〜3（中）。フィルタリングの主キー |
| 3 | 名前 | 氏名 | 通常表示用。名札・名列表等で使用 |
| 4 | ふりがな | 氏名かな | ひらがな表記 |
| 5 | 正式名前 | 正式氏名 | IPAmj明朝用（IVS 付き異体字含む）。公式帳票で使用 |
| 6 | 正式名前ふりがな | 正式氏名かな | 正式ふりがな |
| 7 | 性別 | 性別 | 修了台帳、学級編成用個票で使用 |
| 8 | 生年月日 | 生年月日 | 台帳・個票・家庭調査票で使用。日付型 |
| 9 | 外国籍 | 外国籍 | 学級編成用個票で参考情報 |
| 10 | 郵便番号 | 郵便番号 | 家庭調査票で使用 |
| 11 | 都道府県 | 都道府県 | 住所構成要素① |
| 12 | 市区町村 | 市区町村 | 住所構成要素② |
| 13 | 町番地 | 町番地 | 住所構成要素③ |
| 14 | アパート/マンション名 | 建物名 | 住所構成要素④ |
| 15 | 電話番号1 | 電話番号1 | 主電話番号 |
| 16 | 電話番号2 | 電話番号2 | 副電話番号 |
| 17 | 電話番号3 | 電話番号3 | 副電話番号 |
| 18 | FAX番号 | FAX番号 | 参考情報 |
| 19 | 出身校 | 出身校 | 卒業台帳等で参考 |
| 20〜22 | 出身校住所 / 在籍開始日 / 終了日 | — | 帳票では直接使用しない |
| 23 | 入学日 | 入学日 | 参考情報 |
| 24〜29 | 転入前学校 / 住所 / 在籍日 / 転入日 / 事由 | 転入日 等 | 修了台帳の「学校をまった日」に転入日を使用可能 |
| 30〜35 | 編入前学校 / 住所 / 在籍日 / 編入日 / 事由 | — | 参考情報 |
| 36 | 保護者1　続柄 | 保護者続柄 | 家庭調査票で使用 |
| 37 | 保護者1　名前 | 保護者名 | 家庭調査票・台帳で使用 |
| 38 | 保護者1　名前ふりがな | 保護者名かな | 家庭調査票で使用 |
| 39 | 保護者1　正式名前 | 保護者正式名 | IPAmj明朝用。公式帳票の保護者名に使用 |
| 40 | 保護者1　正式名前ふりがな | 保護者正式名かな | 公式帳票のふりがなに使用 |
| 41 | 保護者1　郵便番号 | 保護者郵便番号 | 保護者住所が児童と異なる場合 |
| 42〜44 | 保護者1　都道府県/市区町村/町番地 | 保護者住所 | 結合して使用 |
| 45 | 保護者1　アパート/マンション名 | 保護者建物名 | 保護者住所の一部 |
| 46 | 保護者1　電話番号1 | 保護者電話1 | 保護者連絡先 |
| 47〜48 | 保護者1　電話番号2/3 | 保護者電話2/3 | 保護者副連絡先 |
| 49 | 保護者1　FAX番号 | 保護者FAX | 参考情報 |
| 50 | 保護者1　緊急連絡先 | 緊急連絡先 | 家庭調査票の緊急連絡先欄に使用 |

### 2.2 カラムマッピング辞書 — 実装コード

```python
# core/mapper.py

# C4th確定ヘッダー → 内部論理名（完全一致マップ）
EXACT_MAP = {
    '生徒コード': '生徒コード',
    '学年': '学年',
    '名前': '氏名',
    'ふりがな': '氏名かな',
    '正式名前': '正式氏名',
    '正式名前ふりがな': '正式氏名かな',
    '性別': '性別',
    '生年月日': '生年月日',
    '外国籍': '外国籍',
    '郵便番号': '郵便番号',
    '都道府県': '都道府県',
    '市区町村': '市区町村',
    '町番地': '町番地',
    'アパート/マンション名': '建物名',
    '電話番号1': '電話番号1',
    '電話番号2': '電話番号2',
    '電話番号3': '電話番号3',
    'FAX番号': 'FAX番号',
    '出身校': '出身校',
    '入学日': '入学日',
    '転入日': '転入日',
    '保護者1\u3000続柄': '保護者続柄',
    '保護者1\u3000名前': '保護者名',
    '保護者1\u3000名前ふりがな': '保護者名かな',
    '保護者1\u3000正式名前': '保護者正式名',
    '保護者1\u3000正式名前ふりがな': '保護者正式名かな',
    '保護者1\u3000郵便番号': '保護者郵便番号',
    '保護者1\u3000都道府県': '保護者都道府県',
    '保護者1\u3000市区町村': '保護者市区町村',
    '保護者1\u3000町番地': '保護者町番地',
    '保護者1\u3000アパート/マンション名': '保護者建物名',
    '保護者1\u3000電話番号1': '保護者電話1',
    '保護者1\u3000電話番号2': '保護者電話2',
    '保護者1\u3000電話番号3': '保護者電話3',
    '保護者1\u3000FAX番号': '保護者FAX',
    '保護者1\u3000緊急連絡先': '緊急連絡先',
}

# 追加エイリアス（他校や旧バージョンで表記が揺れる場合用）
COLUMN_ALIASES = {
    '氏名': ['名前', '児童氏名', '生徒氏名', '児童名'],
    '氏名かな': ['ふりがな', 'フリガナ', 'かな', 'カナ'],
    '正式氏名': ['正式名前', '正式氏名'],
    '出席番号': ['番号', '出席番号', 'No', 'NO', '席番'],
    '組': ['組', '学級', 'クラス'],
}
```

> ⚠ **落とし穴**: C4th ヘッダーの「保護者1　続柄」は全角スペース（U+3000）で区切られている。半角スペースとの混在に注意。`normalize_header` で統一処理すること。

### 2.3 ヘッダー正規化と自動マッピング — 実装コード

```python
def normalize_header(s):
    """ヘッダー名を正規化"""
    if not isinstance(s, str):
        return str(s)
    s = s.strip()
    s = s.replace('\u3000', '\u3000')  # 全角スペースを統一
    s = s.replace('　', '\u3000')       # 万が一のエンコーディング差
    return s

def map_columns(df):
    """DataFrame のカラム名を内部論理名にマッピング"""
    mapped = {}
    unmapped = []
    for col in df.columns:
        norm = normalize_header(col)
        if norm in EXACT_MAP:
            mapped[col] = EXACT_MAP[norm]
        else:
            unmapped.append(col)
    df_mapped = df.rename(columns=mapped)
    return df_mapped, unmapped
```

### 2.4 C4th Excel ファイル読み込み — 実装コード

C4th エクスポートファイルはヘッダー行の前にメタ情報行（学校名、出力日等）が含まれる可能性がある。自動検出が必須。

```python
# core/importer.py
import pandas as pd
from openpyxl import load_workbook

def detect_header_row(filepath, max_scan=10):
    """
    ヘッダー行を自動検出する。
    判定基準: 文字列セルが5つ以上連続する最初の行
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    for row_idx, row in enumerate(ws.iter_rows(max_row=max_scan), 1):
        str_count = sum(
            1 for cell in row
            if cell.value is not None and isinstance(cell.value, str)
        )
        if str_count >= 5:
            wb.close()
            return row_idx
    wb.close()
    return 1  # フォールバック: 1行目をヘッダーとする

def import_c4th_excel(filepath):
    """C4th エクスポート Excel を読み込み、pandas の DataFrame として返す"""
    header_row = detect_header_row(filepath)
    df = pd.read_excel(
        filepath,
        header=header_row - 1,  # 0-indexed
        dtype=str,               # 全列を文字列として読み込み（型変換は後で行う）
    )
    df = df.loc[:, df.columns.notna()]  # 空白カラム名を除去
    df = df.dropna(how='all')           # 全空白行を除去
    return df
```

### 2.5 「名前」vs「正式名前」の使い分け

C4th では「名前」（通常表示用）と「正式名前」（異体字を含む正式表記）の 2 系統が存在する。

| 帳票 | 使用するフィールド | 理由 |
|------|------------------|------|
| 名札（全3種） | 氏名（名前） | 名札は日常使用のため、通常表示で十分 |
| 掲示用名列表 | 氏名（名前） | 教室掲示用 |
| 調べ表 | 氏名（名前） | 日常使用 |
| 修了台帳 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 卒業台帳 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 家庭調査票 | 正式氏名（正式名前） | 公式書類のため正式表記が必須 |
| 学級編成用個票 | 正式氏名（正式名前） | 引継ぎ書類のため正式表記 |

```python
def resolve_name_fields(data_row, use_formal):
    """テンプレートの use_formal_name フラグに基づき氏名フィールドを選択"""
    if use_formal:
        return {
            '表示氏名': data_row.get('正式氏名', data_row.get('氏名', '')),
            '表示氏名かな': data_row.get('正式氏名かな', data_row.get('氏名かな', '')),
            '表示保護者名': data_row.get('保護者正式名', data_row.get('保護者名', '')),
        }
    else:
        return {
            '表示氏名': data_row.get('氏名', ''),
            '表示氏名かな': data_row.get('氏名かな', ''),
            '表示保護者名': data_row.get('保護者名', ''),
        }
# 注意: 正式名前が空の場合は通常の名前にフォールバック
```

### 2.6 住所結合処理

```python
# utils/address.py
def build_address(row):
    """都道府県 + 市区町村 + 町番地 + 建物名 を結合"""
    parts = [
        str(row.get('都道府県', '')),
        str(row.get('市区町村', '')),
        str(row.get('町番地', '')),
        str(row.get('建物名', '')),
    ]
    return ''.join(p for p in parts if p and p != 'nan')
```

### 2.7 必須追加カラム（学年・組・出席番号）

**【最重要】** C4th エクスポートには「学年」は含まれるが、「組」と「出席番号」は含まれない。この 2 つは全テンプレートで必要な基本情報であるため、ファイル読込直後に必ず入力させる必須ステップとして実装する。「帳票生成」ボタンは、組・出席番号が未設定の場合は押せない（disabled）とする。

| 必須カラム | C4th に含まれるか | 補完方法 | UI での操作 |
|-----------|-----------------|---------|-----------|
| 学年 | ◎ 含まれる | C4th から自動読込（手動編集可） | 読込後に自動セット。ドロップダウンで変更可 |
| 組 | × 含まれない | 教員が必ず指定 | 必須入力ドロップダウン（1〜6組 + 手動入力）。未設定時は赤枠警告 |
| 出席番号 | × 含まれない | 名簿順で自動連番（デフォルト）or 手動 | 「自動連番」ボタン + 個別セル編集可 |

### 2.8 必須カラム入力フロー

```
① C4th ファイル読込完了
② 「必須情報入力パネル」が自動表示される：
   ┌─────────────────────────────────┐
   │  📋 必須情報を設定してください      │
   │                                   │
   │  学年：[▼ 1年]（C4th から自動セット）│
   │  組　：[▼ 未設定 ⚠️]  ← 赤枠      │
   │  出席番号：[🔢 自動連番] [手動入力]  │
   │                                   │
   │       [✅ 確定して進む]             │
   └─────────────────────────────────┘
③ 教員が「組」をドロップダウンで選択（例：1組）
④ 「自動連番」をクリック → 名簿順に 1,2,3... が自動付与
⑤ 「確定して進む」ボタンで必須入力完了
⑥ テンプレート選択・生成ボタンが有効化される
```

**複数クラス対応**: 1 つの C4th ファイルに複数クラスのデータが混在する場合は、学年列の値でグループ化し、各グループに対して組・出席番号を設定する画面を提供する。または、ファイルを学級ごとに分けて読み込む運用とする（config.json の設定で切替可能）。

### 2.9 任意追加カラム（手動カラム追加機能）

教員が GUI 上で任意のカラムを追加できる機能。

操作フロー:
1. 「列を追加」ボタンをクリック
2. ダイアログで列名を入力（例：「担任名」「進学先」）
3. 新しい列がデータテーブルの右端に追加
4. 各セルをダブルクリックして直接値を入力可能
5. 「一括入力」ボタンで全行に同じ値を一括セット
6. 「連番入力」ボタンで開始番号を指定し連番を自動入力
7. 追加した列データはセッション中保持、帳票生成時に使用
8. 列データを保存（次回同じファイルを開いた際に自動復元）

---

## 第3章 GUI 設計

### 3.1 メインウィンドウ

| 項目 | 仕様 |
|------|------|
| ウィンドウタイトル | 名簿帳票ツール v1.0 |
| 初期サイズ | 900 x 700 px |
| 最小サイズ | 800 x 600 px |
| テーマ | CustomTkinter "blue" テーマ、ライトモード |
| フォント | メイリオ 11pt（UI 全体） |

### 3.2 左パネル（操作エリア、幅 300px 固定）

**【セクション 1: 名簿読込】**
- 「名簿ファイルを選択」ボタン（CTkButton）
- 選択中ファイル名表示
- 読込件数表示（例：「45 名 読み込み完了」）

**【セクション 2: 必須情報入力 ⚠️】** ← ファイル読込後に自動展開
- 学年ドロップダウン（C4th から自動セット、編集可）
- 組ドロップダウン（必須。未設定時は赤枠 ⚠️ 警告表示）、選択肢：1〜6 組 + 手動入力
- 出席番号設定：「自動連番」ボタン → 名簿順に 1,2,3... を付与、右パネルのプレビュー上で個別編集も可能
- 「✅ 確定して進む」ボタン
- ※ 組・出席番号が未設定の場合、セクション 3 以降はグレーアウト

**【セクション 3: テンプレート選択】** ← 必須情報確定後に有効化
- テンプレートリスト（CTkRadioButton 群）
- アイコン + ラベル + 簡易説明

**【セクション 4: 出力オプション】**
- 年度（自動設定、編集可）
- 担任名（手動入力、前回値を記憶）
- 学校名（config.json から読込、編集可）

**【セクション 5: アクションボタン】** ← 必須情報確定後に有効化
- 「Excel 生成」ボタン（大きめ、アクセントカラー）
- 「直接印刷」ボタン（xlwings 利用可能時のみ）
- 進捗バー

**【セクション 6: システム】**
- 「更新を確認」ボタン（小さめ、グレー）
- バージョン表示

### 3.3 右パネル（プレビュー + カラム編集エリア）

- 読込データのテーブルプレビュー（スクロール可能）
- ツールバー：「列を追加」「一括入力」「連番入力」「列を削除」ボタン
- セルをダブルクリックで直接編集可能
- ステータスバー（下部）：状態メッセージ + 更新通知

### 3.4 GUI 実装コード — アプリケーションクラス

```python
# gui/app.py
import customtkinter as ctk
import threading

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('名簿帳票ツール v1.0')
        self.geometry('900x700')
        self.minsize(800, 600)
        ctk.set_appearance_mode('light')
        ctk.set_default_color_theme('blue')

        # データ状態
        self.df = None           # 読込済み DataFrame
        self.df_mapped = None    # マッピング済み DataFrame
        self.mandatory_ok = False  # 必須カラム入力完了フラグ

        # レイアウト: 2カラム
        self.grid_columnconfigure(0, weight=0, minsize=300)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 左パネル
        self.left_panel = LeftPanel(self, width=300)
        self.left_panel.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

        # 右パネル
        self.right_panel = RightPanel(self)
        self.right_panel.grid(row=0, column=1, sticky='nsew', padx=5, pady=5)

        # バックグラウンド更新チェック
        threading.Thread(target=self._check_update, daemon=True).start()

    def _check_update(self):
        from core.updater import check_for_updates
        from core.config import load_config
        config = load_config()
        update_info = check_for_updates(config)
        if update_info:
            # UIスレッドでダイアログ表示
            self.after(0, lambda: self._show_update_dialog(update_info))
```

> ⚠ **落とし穴**: CustomTkinter の UI は必ずメインスレッドで操作すること。バックグラウンドスレッドから UI を直接更新するとクラッシュする。`self.after(0, callback)` を使ってメインスレッドに処理を委譲する。

### 3.5 必須カラム入力パネル — 状態管理の実装コード

```python
class MandatoryInputPanel(ctk.CTkFrame):
    """組・出席番号の必須入力パネル"""
    def __init__(self, master, on_confirm):
        super().__init__(master)
        self.on_confirm = on_confirm
        self.is_confirmed = False

        # 組ドロップダウン
        self.kumi_var = ctk.StringVar(value='未選択')
        self.kumi_label = ctk.CTkLabel(self, text='組 ⚠️', text_color='red')
        self.kumi_dropdown = ctk.CTkComboBox(
            self, values=['1', '2', '3', '4', '5', '6'],
            variable=self.kumi_var,
            command=self._validate
        )

        # 出席番号ボタン
        self.shusseki_btn = ctk.CTkButton(
            self, text='🔢 自動連番',
            command=self._auto_number
        )

        # 確定ボタン（初期状態は disabled）
        self.confirm_btn = ctk.CTkButton(
            self, text='✅ 確定して進む',
            state='disabled',
            command=self._confirm
        )

    def _validate(self, *args):
        """組が選択されているか検証"""
        if self.kumi_var.get() != '未選択':
            self.kumi_label.configure(text='組', text_color='black')
            self.confirm_btn.configure(state='normal')
        else:
            self.kumi_label.configure(text='組 ⚠️', text_color='red')
            self.confirm_btn.configure(state='disabled')
```

---

## 第4章 テンプレート仕様（全 9 種）

**共通仕様**: テンプレートは全て openpyxl でプログラム生成する（既存テンプレートなし）。プレースホルダー = `{{論理名}}` 形式。公式帳票では `use_formal_name=true` とし、正式名前/正式名前ふりがなを使用する。

### 4.0 差込処理の 4 タイプ

| タイプ | 説明 | 対象テンプレート |
|--------|------|-----------------|
| grid（グリッド配置型） | 1 ページに複数名を格子状配置 | 名札 3 種 |
| list（リスト展開型） | データ行を児童数分コピー展開 | 掲示用名列表、修了台帳、卒業台帳、調べ表 |
| individual（個票複製型） | 1 ページ 1 児童のシートを複製 | 家庭調査票、学級編成用個票 |
| replace（単純置換型） | プレースホルダーを値で置換 | 全タイプの基本処理 |

### 4.0.1 template_registry.py

```python
TEMPLATES = {
    "名札_通常": {
        "file": "名札_通常.xlsx",
        "type": "grid",
        "cards_per_page": 10,
        "grid_cols": 2, "grid_rows": 5,
        "use_formal_name": False,
        "required_columns": ["氏名", "氏名かな"],
        "mandatory_columns": ["組", "出席番号"],
        "icon": "🏷", "description": "机用名札（装飾なし）10枚/ページ"
    },
    "修了台帳": {
        "file": "修了台帳.xlsx",
        "type": "list",
        "orientation": "landscape",
        "use_formal_name": True,
        "required_columns": ["正式氏名", "正式氏名かな", "性別", "生年月日"],
        "optional_columns": ["保護者正式名", "住所", "電話番号1", "転入日"],
        "icon": "📋", "description": "修了台帳（横A4・一覧表）"
    },
    "家庭調査票": {
        "file": "家庭調査票.xlsx",
        "type": "individual",
        "use_formal_name": True,
        "required_columns": ["正式氏名", "正式氏名かな", "生年月日", "性別"],
        "optional_columns": ["郵便番号", "住所", "電話番号1", "保護者正式名",
                             "保護者続柄", "緊急連絡先"],
        "icon": "📄", "description": "家庭調査票（1名/ページ個票）"
    },
    # ... 他テンプレートも同様
}
```

### 4.0.2 共通レイアウト定数

| 定数名 | 値 | 説明 |
|--------|-----|------|
| PINK_BG | #FFE0E0 | データセルの薄ピンク背景色 |
| FONT_NAME_L | IPAmj明朝, 36pt | 名札の氏名用 |
| FONT_NAME_M | IPAmj明朝, 14pt | 名列表の氏名用 |
| FONT_NAME_S | IPAmj明朝, 10pt | 台帳の氏名用 |
| FONT_LABEL | IPAmj明朝, 8pt | ラベル用 |
| FONT_TITLE | IPAmj明朝, 28pt, bold | タイトル用 |

---

### 4.1 名札_通常（机用名札・装飾なし）

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 横置き |
| レイアウト | 2 列 x 5 行 = 1 ページ 10 名 |
| use_formal_name | false |
| 各名札 | 上部小文字ふりがな（丸付き）、下部大文字氏名（丸付き） |
| 丸の実装 | 各文字を個別セルに配置、太枠線で正方形に近い外観を作る |
| フォント | IPAmj明朝 36pt（氏名）/ 12pt（ふりがな） |
| 差込 | `{{氏名}}`, `{{氏名かな}}` のみ |
| 特殊処理 | 姓名分割→各文字を個別セルに配置。2〜4 文字対応で自動調整 |

> ⚠ **落とし穴**: openpyxl は Excel のオートシェイプ（円、四角形等の図形描画）をサポートしていない。「丸に文字」は図形ではなくセル罫線で代替する。

**実装コード:**

```python
from openpyxl.styles import Border, Side, Alignment, Font

THICK = Side(style='medium', color='000000')
CHAR_BORDER = Border(top=THICK, bottom=THICK, left=THICK, right=THICK)
CENTER = Alignment(horizontal='center', vertical='center')

def write_name_tag(ws, start_row, start_col, name, furigana):
    """1名分の名札を指定位置に書き込む"""
    # 姓と名を分割
    parts = name.split() if ' ' in name else [name[:len(name)//2], name[len(name)//2:]]
    sei, mei = parts[0], parts[1] if len(parts) > 1 else ''

    # 各文字を個別セルに配置（氏名）
    col = start_col
    for char in sei + mei:
        cell = ws.cell(row=start_row + 2, column=col)
        cell.value = char
        cell.font = Font(name='IPAmj明朝', size=36)
        cell.alignment = CENTER
        cell.border = CHAR_BORDER
        ws.column_dimensions[cell.column_letter].width = 7
        col += 1
    ws.row_dimensions[start_row + 2].height = 50  # 正方形に近い高さ
```

### 4.2 名札_装飾あり（花柄ボーダー付き）

| 項目 | 仕様 |
|------|------|
| 基本 | 名札_通常と同一 |
| 追加 | 各名札の上下に花柄ボーダー画像を配置 |
| 画像 | flower_border.png を assets/ に同梱 |

**画像挿入コード:**

```python
from openpyxl.drawing.image import Image

def add_border_image(ws, img_path, cell_anchor, width_px, height_px):
    """花柄ボーダー等の装飾画像をセルに配置"""
    img = Image(img_path)
    img.width = width_px
    img.height = height_px
    img.anchor = cell_anchor  # 例: 'A1'
    ws.add_image(img)
    # 注意: 画像はセルに「固定」されるのではなく
    # セル位置を基準に「フローティング」で配置される
    # → 列幅・行高の変更で位置がずれる可能性あり
```

> ⚠ **落とし穴**: Pillow 依存。openpyxl で画像を扱うには Pillow が必要。requirements.txt に含めること。

### 4.3 名札_1 年生用（縦長短冊型）

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 縦置き |
| レイアウト | 8 列 = 1 ページ 8 名 |
| 背景色 | 薄ピンク #FFE0E0 |
| 上部 | 「番号（番越し）」= 出席番号 |
| 名前部分 | ふりがなを縦書き大きく表示（ひらがなのみ） |
| フォント | IPAmj明朝 48pt、縦書き（textRotation=255） |
| use_formal_name | false |

**縦書き実装コード:**

```python
from openpyxl.styles import Alignment

VERTICAL = Alignment(
    textRotation=255,      # 255 = 縦書き（tategaki）
    horizontal='center',
    vertical='center',
)
cell.alignment = VERTICAL
cell.value = 'たろう'  # 縦書きで表示される
```

> ⚠ **落とし穴**: `textRotation=255` は openpyxl 2.5 以降でサポート。0〜180 は回転角度、255 のみが特別に「縦書き」を意味する。

### 4.4 掲示用名列表

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 縦置き |
| レイアウト | 2 列名簿（左: No.1〜20、右: No.21〜40） |
| タイトル | 「学年」「学級」を中央に大きく表示 |
| データ行 | 出席番号 + 氏名（ふりがな小文字付き） |
| 背景色 | 薄ピンク #FFE0E0 |
| use_formal_name | false |
| 差込 | `{{出席番号}}`, `{{氏名}}`, `{{氏名かな}}`, `{{学年}}`, `{{組}}` |

### 4.5 調べ表

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 縦置き |
| レイアウト | 6 列 x 10 行グリッド（最大 60 名） |
| タイトル | 左上空白（手書き調査項目名）、右上「調べ/期限」斜線セル |
| サブタイトル | 「学年 学級」左、「担任：担任名」右 |
| 各セル | 上部小文字で番号(ふりがな)、中央に氏名、下部空白（記入欄） |
| use_formal_name | false |
| 差込 | `{{出席番号}}`, `{{氏名}}`, `{{氏名かな}}`, `{{学年}}`, `{{組}}`, `{{担任名}}` |

### 4.6 修了台帳

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 横置き |
| タイトル | 「年度（和暦）学年学級 修了台帳」+「担任：担任名」 |
| ヘッダー | 学校をまった日|番|氏名|ふりがな|性別|生年月日|保護者名|住所|電話番号|在校兄弟 |
| use_formal_name | true（正式名前を使用） |
| 住所 | 都道府県+市区町村+町番地+建物名 を結合して 1 列に表示 |
| 電話 | 電話番号 1 を使用 |
| 保護者 | 保護者正式名 を使用 |
| 学校をまった日 | 転入日があればその値、なければ空白 |
| フォント | IPAmj明朝 8pt（列が多いため小さめ） |

### 4.7 卒業台帳

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 横置き |
| タイトル | 「年度（和暦）卒業生名簿」+「年度（西暦）3月31日卒業 学校名」 |
| ヘッダー | 証書番号|氏名|生年月日|住所|保護者名|進学先|担任名 |
| use_formal_name | true |
| 証書番号 | 手動追加カラム。なければ config.json の開始番号から自動連番 |
| 進学先 | 手動追加カラム。なければ空白 |

### 4.8 家庭調査票

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 縦置き |
| レイアウト | 1 ページ 1 児童の個票（最複雑テンプレート） |
| use_formal_name | true |
| 差込フィールド | 正式氏名、正式氏名かな、生年月日、性別、郵便番号、住所（結合）、電話番号 1、保護者正式名、保護者正式名かな、保護者続柄、緊急連絡先、在校兄弟 |
| 空白セクション | 家族構成、保健調査、児童の生活、自由記入、公開レベル、緊急時引渡し欄 |
| 保護者住所 | 児童住所と同じ場合は「同上」、異なる場合は保護者住所を表示 |

### 4.9 学級編成用個票

| 項目 | 仕様 |
|------|------|
| 用紙 | A4 縦置き |
| レイアウト | 1 ページ 1 児童の個票 |
| use_formal_name | true |
| 差込フィールド | 正式氏名、正式氏名かな、性別、生年月日、住所（結合）、電話番号 1、外国籍 |
| 大部分空白 | 学力、行動特性、要配慮、問題行動、欠席状況、引継ぎ事項すべて空白（手書き用） |

### 4.10 個票複製時の画像対応パターン

> ⚠ **落とし穴**: openpyxl の `copy_worksheet` は画像をコピーしない。

| 項目 | コピーされるか |
|------|--------------|
| セルの値 | ✅ |
| セルのスタイル（フォント・色・罫線） | ✅ |
| 列幅・行高 | ✅ |
| 結合セル | ✅ |
| 余白・印刷設定 | ✅ |
| 画像（Image） | ❌ 手動で再挿入が必要 |
| チャート | ❌ 本プロジェクトでは不使用 |
| データバリデーション | ✅ |
| 条件付き書式 | ✅ |

```python
def copy_sheet_with_images(wb, source_ws, new_title):
    """画像も含めたワークシート複製"""
    target_ws = wb.copy_worksheet(source_ws)
    target_ws.title = new_title
    for img in source_ws._images:
        from copy import copy
        from openpyxl.drawing.image import Image
        new_img = Image(img.ref)
        new_img.anchor = copy(img.anchor)
        new_img.width = img.width
        new_img.height = img.height
        target_ws.add_image(new_img)
    return target_ws
```

---

## 第5章 コアロジック（generator.py）

### 5.1 ジェネレータークラス構造

```python
# core/generator.py
from abc import ABC, abstractmethod
from openpyxl import load_workbook
from utils.font_helper import apply_font
from utils.address import build_address
from templates.template_registry import TEMPLATES

class BaseGenerator(ABC):
    def __init__(self, template_path, output_path, data, options):
        self.template_path = template_path
        self.output_path = output_path
        self.data = data           # pandas DataFrame
        self.options = options      # dict: fiscal_year, teacher_name, school_name等
        self.wb = None

    def generate(self):
        self.wb = load_workbook(self.template_path)
        self._populate()            # サブクラスで実装
        self._apply_font_all()
        self._apply_print_settings()
        self.wb.save(self.output_path)
        return self.output_path

    def _apply_font_all(self):
        for ws in self.wb.worksheets:
            apply_font(ws)

    @abstractmethod
    def _populate(self):
        pass

class GridGenerator(BaseGenerator):
    """名札系: 1ページにN名をグリッド配置"""
    def _populate(self):
        meta = self._get_meta()
        cards_per_page = meta['cards_per_page']
        # テンプレートの最初のシートから1カード分のレイアウトを取得
        # 児童数分のカードをページに配置
        ...

class ListGenerator(BaseGenerator):
    """名列表・台帳系: データ行を児童数分コピー"""
    def _populate(self):
        ws = self.wb.active
        # テンプレート内のプレースホルダー行を特定
        # 児童データで行を展開
        ...

class IndividualGenerator(BaseGenerator):
    """家庭調査票・個票系: 1名/シートで複製"""
    def _populate(self):
        template_ws = self.wb.active
        for i, (_, row) in enumerate(self.data.iterrows()):
            if i == 0:
                ws = template_ws
            else:
                ws = self.wb.copy_worksheet(template_ws)
            ws.title = f'{int(row.get("出席番号", i+1)):02d}_{row["氏名"]}'
            self._fill_placeholders(ws, row)
        ...

def create_generator(template_name, output_path, data, options):
    """ファクトリー関数"""
    meta = TEMPLATES[template_name]
    template_path = os.path.join(options['template_dir'], meta['file'])
    gen_type = meta['type']
    if gen_type == 'grid':
        return GridGenerator(template_path, output_path, data, options)
    elif gen_type == 'list':
        return ListGenerator(template_path, output_path, data, options)
    elif gen_type == 'individual':
        return IndividualGenerator(template_path, output_path, data, options)
    else:
        raise ValueError(f'Unknown generator type: {gen_type}')
```

### 5.2 プレースホルダー置換の共通処理

```python
import re
import pandas as pd

PLACEHOLDER_RE = re.compile(r'\{\{(.+?)\}\}')

def fill_placeholders(ws, data_row, options=None):
    """ワークシート内の全プレースホルダーをデータで置換"""
    options = options or {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and '{{' in cell.value:
                def replacer(match):
                    key = match.group(1)
                    # 特殊キー
                    if key == '年度':
                        return str(options.get('fiscal_year', ''))
                    if key == '年度和暦':
                        from utils.wareki import to_wareki
                        return to_wareki(options.get('fiscal_year', 2025))
                    if key == '学校名':
                        return options.get('school_name', '')
                    if key == '担任名':
                        return options.get('teacher_name', '')
                    if key == '住所':
                        from utils.address import build_address
                        return build_address(data_row)
                    # 通常のカラム参照
                    val = data_row.get(key, '')
                    return '' if pd.isna(val) else str(val)
                new_value = PLACEHOLDER_RE.sub(replacer, cell.value)
                cell.value = new_value
```

### 5.3 font_helper.py

```python
# utils/font_helper.py
from openpyxl.styles import Font

def apply_font(ws, font_name='IPAmj明朝'):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cur = cell.font
                cell.font = Font(
                    name=font_name, size=cur.size,
                    bold=cur.bold, italic=cur.italic, color=cur.color
                )
```

### 5.4 印刷設定の完全な指定方法

```python
from openpyxl.worksheet.page import PageMargins

def setup_print(ws, orientation='portrait', paper_size=9):
    """
    印刷設定を適用
    paper_size: 9=A4, 1=Letter
    orientation: 'portrait' or 'landscape'
    """
    ws.page_setup.paperSize = paper_size
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0=自動

    ws.page_margins = PageMargins(
        left=0.39,    # 10mm ≈ 0.39 inch
        right=0.39,
        top=0.39,
        bottom=0.39,
        header=0.2,
        footer=0.2,
    )
    ws.print_options.horizontalCentered = True
```

### 5.5 openpyxl スタイル定義の正しい書き方

> ⚠ **落とし穴**: openpyxl のスタイルオブジェクト（Font, Alignment, Border 等）はイミュータブル（不変）。一度セルに設定した後、個別プロパティを変更することはできない。

```python
# ❌ 間違い（エラーになる）
cell.font.bold = True
cell.font.name = 'IPAmj明朝'

# ✅ 正しい
cell.font = Font(name='IPAmj明朝', size=12, bold=True)

# ✅ 既存フォントのプロパティを保持しつつ変更する場合
from copy import copy
new_font = copy(cell.font)
cell.font = Font(
    name='IPAmj明朝',
    size=new_font.size,
    bold=new_font.bold,
    italic=new_font.italic,
    color=new_font.color,
)
```

### 5.6 結合セル（MergedCell）への書き込み制限

> ⚠ **落とし穴**: 結合セルの左上以外に値を書き込むと `AttributeError: 'MergedCell' object attribute 'value' is read-only` が発生する。

```python
# ❌ エラーになるパターン
ws.merge_cells('B2:D4')
ws['C3'].value = 'テスト'  # AttributeError!

# ✅ 正しいパターン（左上セルに書き込む）
ws.merge_cells('B2:D4')
ws['B2'].value = 'テスト'
ws['B2'].alignment = Alignment(horizontal='center', vertical='center')

# ✅ プレースホルダー置換時の安全チェック
from openpyxl.cell.cell import MergedCell

def safe_set_value(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = ws.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col
                )
                top_left.value = value
                return
    else:
        cell.value = value
```

### 5.7 結合セルの罫線

```python
from openpyxl.styles import Border, Side

def set_merged_cell_border(ws, cell_range, border):
    """結合セル範囲の外枠に罫線を確実に適用"""
    rows = list(ws[cell_range])
    for row in rows:
        for cell in row:
            new_border = Border(
                top=border.top if cell.row == rows[0][0].row else None,
                bottom=border.bottom if cell.row == rows[-1][0].row else None,
                left=border.left if cell.column == rows[0][0].column else None,
                right=border.right if cell.column == rows[0][-1].column else None,
            )
            cell.border = new_border
```

### 5.8 NamedStyle で効率的にスタイル管理

```python
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

def register_styles(wb):
    """アプリ共通の NamedStyle を Workbook に登録"""
    name_tag = NamedStyle(name='名札_氏名')
    name_tag.font = Font(name='IPAmj明朝', size=36)
    name_tag.alignment = Alignment(horizontal='center', vertical='center')
    name_tag.border = Border(
        top=Side('medium'), bottom=Side('medium'),
        left=Side('medium'), right=Side('medium'),
    )
    wb.add_named_style(name_tag)

    header = NamedStyle(name='一覧_ヘッダー')
    header.font = Font(name='Yu Gothic', size=10, bold=True)
    header.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
    header.border = Border(
        top=Side('thin'), bottom=Side('thin'),
        left=Side('thin'), right=Side('thin'),
    )
    wb.add_named_style(header)

    data_row = NamedStyle(name='一覧_データ')
    data_row.font = Font(name='Yu Gothic', size=10)
    data_row.alignment = Alignment(vertical='center')
    data_row.border = Border(
        top=Side('thin'), bottom=Side('thin'),
        left=Side('thin'), right=Side('thin'),
    )
    wb.add_named_style(data_row)

# 使用例
cell.style = '名札_氏名'
```

> ⚠ **落とし穴**: NamedStyle は一度セルに適用した後、スタイル定義を変更しても既適用セルには反映されない。

### 5.9 openpyxl で load_workbook → save 時の画像消失

> ⚠ **落とし穴**: openpyxl で `load_workbook` → `save` すると、元ファイルにあった画像やチャートが消える可能性がある（openpyxl 3.x 系の既知制限）。

**対策**: テンプレートをコピーしてから編集。

```python
import shutil

def generate_from_template(template_path, output_path, data):
    shutil.copy2(template_path, output_path)
    wb = load_workbook(output_path)
    # ... データ差込処理 ...
    wb.save(output_path)
    wb.close()
```

### 5.10 Excel 寸法チートシート

| 項目 | openpyxl 単位 | 1mm の値 | 備考 |
|------|-------------|---------|------|
| 列幅 | 文字数（半角基準） | ≈0.35 | フォントに依存 |
| 行高 | ポイント（pt） | ≈2.83 | 1pt = 1/72 inch |
| 余白 | インチ | ≈0.039 | PageMargins |
| 画像幅/高 | ピクセル（px） | ≈3.78 | 96dpi 基準 |

実用的な列幅の目安:

| 内容 | 列幅（文字数） | 概算 mm |
|------|-------------|--------|
| 1 文字セル（名札の個別文字） | 6〜7 | 約 17mm |
| 出席番号 | 4〜5 | 約 12mm |
| 氏名（通常） | 12〜15 | 約 35〜43mm |
| 氏名（正式・長め） | 18〜20 | 約 52〜58mm |
| 住所 | 35〜45 | 約 100〜130mm |
| 電話番号 | 14〜16 | 約 40〜46mm |

---

## 第6章 Google Drive 自動更新

### 6.1 概要

アプリ本体（exe）およびテンプレートファイルを Google Drive の共有フォルダに配置し、起動時に自動で最新版を確認・ダウンロードする。RYUMA が 1 箇所（Google Drive）を更新するだけで、全校のアプリが自動的に最新化される。

### 6.2 Google Drive 側の構成

「リンクを知っている全員が閲覧可」で共有。

```
Google Drive:
  名簿帳票ツール_配布/
    ├── version.json           # バージョン管理ファイル
    ├── 名簿帳票ツール.exe      # アプリ本体（最新版）
    ├── テンプレート/
    │   ├── 名札_通常.xlsx
    │   ├── ... (全テンプレート)
    └── assets/
        └── flower_border.png
```

### 6.3 version.json の仕様

```json
{
  "app_version": "1.0.2",
  "app_file_id": "1AbCdEfG_googledrivefileid",
  "release_date": "2026-03-15",
  "release_notes": "名札テンプレートの余白を調整",
  "templates": {
    "version": "1.0.1",
    "files": {
      "名札_通常.xlsx": "1XyZaBC_fileid1",
      "名札_装飾あり.xlsx": "1XyZaBC_fileid2",
      "名札_1年生用.xlsx": "1XyZaBC_fileid3",
      "掲示用名列表.xlsx": "1XyZaBC_fileid4",
      "調べ表.xlsx": "1XyZaBC_fileid5",
      "修了台帳.xlsx": "1XyZaBC_fileid6",
      "卒業台帳.xlsx": "1XyZaBC_fileid7",
      "家庭調査票.xlsx": "1XyZaBC_fileid8",
      "学級編成用個票.xlsx": "1XyZaBC_fileid9"
    }
  }
}
```

`app_file_id` / `files` の値: 各ファイルの Google Drive ファイル ID（共有リンク `https://drive.google.com/file/d/{FILE_ID}/view` の `{FILE_ID}` 部分）。

### 6.4 自動更新処理フロー

アプリ起動時に非同期（別スレッド）で実行。メイン画面の表示を妨げないこと。

1. `config.json` から `version_file_id` を読み取る
2. `requests.get()` で version.json をダウンロード（タイムアウト 5 秒）
3. ダウンロード失敗時（オフライン等）→ スキップして通常起動。エラーは表示しない
4. ダウンロード成功時 → ローカルのバージョンと比較
5. a. アプリ本体が更新 → ダイアログ表示 →「はい」で exe ダウンロード＆差し替え＆再起動
   b. テンプレートが更新 → バックグラウンドで差分テンプレートのみダウンロード・上書き → ステータスバーに通知
   c. 両方最新 → 何もしない

### 6.5 Google Drive ダウンロード — 実装コード（API キー不要）

```python
import requests
import os

def download_from_gdrive(file_id, dest_path, progress_cb=None):
    """
    Google Drive からファイルをダウンロード（APIキー不要）
    共有設定が「リンクを知っている全員が閲覧可」であること
    """
    URL = 'https://drive.google.com/uc?export=download'
    session = requests.Session()

    response = session.get(URL, params={'id': file_id}, stream=True)

    # 大きいファイル（>25MB）のウイルススキャン確認ページ対応
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break
    if token:
        response = session.get(
            URL, params={'id': file_id, 'confirm': token}, stream=True
        )

    total = int(response.headers.get('content-length', 0))
    downloaded = 0
    with open(dest_path, 'wb') as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)
                downloaded += len(chunk)
                if progress_cb and total > 0:
                    progress_cb(downloaded / total)
    return dest_path
```

> ⚠ **重要**: Google Drive の「リンクを知っている全員が閲覧可」設定であれば、Google API キーや OAuth 認証は不要。requests ライブラリのみでダウンロード可能。学校のネットワークで Google ドメインへのアクセスが許可されていることが前提。

### 6.6 exe 自己更新バッチファイル — 実装コード

> ⚠ **落とし穴**: Windows では実行中の exe ファイルを上書きできない（ファイルロック）。バッチファイルで待機→置換→再起動する。

```python
import subprocess, sys, os

def generate_update_batch(new_exe_path, current_exe_path):
    """
    自己更新用バッチファイルを生成・実行
    """
    exe_name = os.path.basename(current_exe_path)
    exe_dir = os.path.dirname(current_exe_path)

    bat_content = f'''@echo off
chcp 65001 > nul
timeout /t 2 /nobreak > nul
taskkill /IM "{exe_name}" /F > nul 2>&1
timeout /t 1 /nobreak > nul
del "{current_exe_path}"
move "{new_exe_path}" "{current_exe_path}"
start "" "{current_exe_path}"
del "%~f0"
'''
    bat_path = os.path.join(exe_dir, '_update.bat')
    with open(bat_path, 'w', encoding='utf-8') as f:
        f.write(bat_content)

    subprocess.Popen(
        ['cmd', '/c', bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
        cwd=exe_dir,
    )
    sys.exit(0)
```

> ⚠ **落とし穴**: `chcp 65001` は必須。日本語ファイル名を含むパスをバッチファイルで正しく扱うために UTF-8 コードページに切り替える。これがないと文字化けでファイルが見つからずエラーになる。

### 6.7 updater.py ソースコード構成

```
core/updater.py
  ├── check_for_updates(config) -> UpdateInfo | None
  │     version.json を取得し、ローカルと比較
  ├── download_file(file_id, dest_path, progress_callback)
  │     Google Drive からファイルをダウンロード
  ├── update_app(update_info)
  │     exe 差し替えバッチ生成→再起動
  ├── update_templates(update_info, template_dir)
  │     差分テンプレートのみダウンロード・上書き
  └── UpdateInfo (dataclass)
        app_version, template_version, app_file_id, template_files, release_notes
```

### 6.8 日本語パスの安全な取得

```python
import os, sys

def get_app_dir():
    """アプリケーションの実行ディレクトリを安全に取得"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_output_dir():
    base = get_app_dir()
    out = os.path.join(base, '出力')
    os.makedirs(out, exist_ok=True)
    return out

def get_template_dir():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'テンプレート')
    else:
        return os.path.join(get_app_dir(), 'テンプレート')
```

> ⚠ **落とし穴**: `os.path.join()` は日本語パスを正しく扱えるが、subprocess で bat ファイルを実行する際は `chcp 65001` で UTF-8 モードにすること。

---

## 第7章 ビルド・配布・実装順序・テスト

### 7.1 PyInstaller ビルド（--onedir 必須）

> ⚠ **最重要**: CustomTkinter は .json/.otf などの非 Python データファイルを含むため、`--onefile` では正常に動作しない。必ず `--onedir` を使用する。

**問題の原因**: CustomTkinter はテーマ定義（.json）とフォントファイル（.otf）をパッケージ内に含む。`--onefile` モードはこれらを一時フォルダに展開するが、CustomTkinter は自身のパッケージディレクトリからの相対パスでこれらを探すため、ファイルが見つからずクラッシュする。

**build.spec:**

```python
from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('customtkinter')
datas += [('テンプレート', 'テンプレート')]
datas += [('config.json', '.')]

a = Analysis(
    ['main.py'],
    datas=datas,
    hiddenimports=['openpyxl', 'customtkinter'],
)
pyz = PYZ(a.pure, a.zipped_data)
exe = EXE(pyz, a.scripts, [],
    name='名簿帳票ツール',
    icon='assets/icon.ico',
    console=False,
)
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas,
    name='名簿帳票ツール')
```

ビルドコマンド: `pyinstaller build.spec`

出力は `dist/名簿帳票ツール/` フォルダ。配布時はこのフォルダ全体を共有する。

### 7.2 --onefile を使いたい場合のワークアラウンド（非推奨）

```python
# main.py の冒頭に追記（他の import より前に実行）
import os, sys
if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
    os.chdir(application_path)
    sys.path.insert(0, application_path)
import customtkinter
```

ただし config やテンプレートの読み書きパスが複雑になるため非推奨。

### 7.3 配布フロー

**初回:**
1. ビルド済み exe とテンプレートを Google Drive にアップロード
2. version.json を作成し、各ファイル ID を記載
3. 全ファイルを「リンクを知っている全員が閲覧可」に設定
4. 各校に Google Drive フォルダの URL を共有
5. 教員が exe とテンプレートフォルダをローカルにコピー
6. 以降はアプリの自動更新機能で最新化

**2 回目以降:**
1. RYUMA が exe またはテンプレートを修正
2. Google Drive の該当ファイルを差し替え
3. version.json のバージョン番号を更新
4. 各校の教員がアプリを起動すると自動で更新通知 → ワンクリックで更新完了

### 7.4 プロジェクト作成手順

```bash
# 1. 仮想環境作成
python -m venv venv
venv\Scripts\activate

# 2. 依存インストール
pip install -r requirements.txt

# 3. 開発実行
python main.py

# 4. テンプレート生成（初回のみ）
python -m templates.generators.generate_all

# 5. ビルド
pyinstaller build.spec

# 6. 出力確認
# dist/名簿帳票ツール/ フォルダが生成される
```

### 7.5 推奨実装順序（ファイル単位、依存関係順）

| 順 | ファイル | 内容 | 依存 | テスト方法 |
|----|--------|------|------|-----------|
| 1 | utils/wareki.py | 和暦変換 | なし | pytest: 2019→令和1年、1989→平成1年 |
| 2 | utils/address.py | 住所結合 | なし | pytest: 4フィールド結合、空値除外 |
| 3 | utils/font_helper.py | フォント設定 | openpyxl | Excel で開いてフォント名確認 |
| 4 | core/config.py | 設定管理 | json | config.json 読み書き |
| 5 | core/mapper.py | カラムマッピング | なし | EXACT_MAP の完全一致テスト |
| 6 | core/importer.py | Excel 読込 | mapper,pandas | ダミーデータで読込テスト |
| 7 | templates/template_registry.py | テンプレートメタ | なし | 辞書参照テスト |
| 8 | templates/generators/gen_meireihyo.py | 名列表テンプレート生成 | openpyxl | Excel で開いてレイアウト確認 |
| 9 | core/generator.py (ListGenerator) | リスト型差込 | importer | 名列表に 35 名データ差込 |
| 10 | gui/app.py + frames/ | メイン GUI | customtkinter | ウィンドウ表示・操作確認 |
| 11 | 統合テスト | 読込→差込→出力 | 全部 | ダミーデータで名列表出力 |
| 12 | 名札テンプレート 3 種 | GridGenerator | generator | 名札出力確認 |
| 13 | 台帳テンプレート 2 種 | ListGenerator 拡張 | generator | 台帳出力確認 |
| 14 | core/updater.py | 自動更新 | requests | ローカル HTTP サーバでテスト |
| 15 | 個票テンプレート 2 種 | IndividualGenerator | generator | 個票出力確認 |
| 16 | build.spec + ビルド | exe 化 | pyinstaller | 別 PC で起動確認 |

### 7.6 開発フェーズとスケジュール

**Phase 1（MVP）: 3 月第 1〜2 週**
- 名札 3 種 + 掲示用名列表
- 組・出席番号の必須入力対応
- 任意カラム追加対応

**Phase 1.5: 3 月第 3 週（パイロット）**
- 天久小の実データでテスト

**Phase 2: 3 月第 3〜4 週**
- 修了台帳 + 卒業台帳 + 調べ表 + 自動更新機能

**Phase 3: 4 月以降**
- 全 9 テンプレート対応の完成版

### 7.7 テスト構成

```
tests/
  ├── test_wareki.py         # 和暦変換のユニットテスト
  ├── test_address.py        # 住所結合のユニットテスト
  ├── test_mapper.py         # カラムマッピングのテスト
  ├── test_importer.py       # Excel 読込テスト
  ├── test_generator.py      # 各ジェネレーターの出力テスト
  ├── test_updater.py        # 更新チェックのモックテスト
  ├── fixtures/
  │   ├── dummy_c4th.xlsx    # ダミー C4th エクスポート
  │   └── expected/          # 期待出力 Excel
  └── generate_dummy.py      # テスト用ダミーデータ生成
```

### 7.8 テストケース一覧

| テスト対象 | ケース | 検証内容 |
|-----------|--------|---------|
| wareki.py | 2019→令和1年 | 西暦から和暦への正しい変換 |
| wareki.py | 1989→平成1年 | 昭和/平成の境界 |
| address.py | 4 フィールド全入力 | 結合結果が「沖縄県那覇市...」 |
| address.py | 建物名が NaN | NaN/空値がスキップされる |
| mapper.py | 全 50 カラム | 全カラムが正しくマッピング |
| mapper.py | 全角スペース混在 | 「保護者1　名前」が正規化される |
| importer.py | ヘッダー行検出 | メタ行がある場合も正しく検出 |
| importer.py | dtype=str 確認 | 生年月日が文字列のまま |
| GridGenerator | 35 名→名札 | 1 ページ 8 枚 × 5 ページの出力 |
| ListGenerator | 35 名→名列表 | 全名がリストに含まれる |
| IndividualGenerator | 35 名→個票 | 35 シートが生成される |
| MandatoryInput | 組未選択 | 生成ボタンが disabled |
| MandatoryInput | 連番生成 | 出席番号が 1〜35 連番 |

### 7.9 手動確認チェックリスト（印刷テスト）

- □ A4 用紙に印刷して文字が切れていないか
- □ 名札の文字サイズが適切か
- □ 名列表の罫線が全て印刷されるか
- □ 正式名前の IVS 文字（異体字）が正しく表示されるか
- □ IPAmj明朝フォントがインストールされていない PC での表示
- □ ふりがなの文字サイズが氏名より十分小さいか
- □ ページ余白が適切か（端が切れない）
- □ 複数ページに跨がるデータでページ区切りが正しいか
- □ Excel 上で「印刷プレビュー」で確認

### 7.10 テスト用ダミーデータ生成

```python
# tests/generate_dummy.py
import pandas as pd
import random

LAST_NAMES = ['山田', '田中', '鈴木', '佐藤', '高橋', '伊藤', '渡辺', '中村', '小林', '加藤']
FIRST_NAMES_M = ['太郎', '一郎', '健太', '翔', '大輔', '拓也', '修', '隆', '誠', '光']
FIRST_NAMES_F = ['花子', '美咲', '陽菜', '結衣', 'さくら', '凛', '葵', '彩', '楓', '愛']
FIRST_KANA_M = ['たろう','いちろう','けんた','しょう','だいすけ','たくや','おさむ','たかし','まこと','ひかる']
FIRST_KANA_F = ['はなこ','みさき','ひな','ゆい','さくら','りん','あおい','あや','かえで','あい']

def generate_dummy(n=35, grade=1):
    rows = []
    for i in range(n):
        sex = random.choice(['男', '女'])
        sei = random.choice(LAST_NAMES)
        if sex == '男':
            idx = random.randint(0, len(FIRST_NAMES_M)-1)
            mei, mei_k = FIRST_NAMES_M[idx], FIRST_KANA_M[idx]
        else:
            idx = random.randint(0, len(FIRST_NAMES_F)-1)
            mei, mei_k = FIRST_NAMES_F[idx], FIRST_KANA_F[idx]
        rows.append({
            '生徒コード': f'S{grade:02d}{i+1:03d}',
            '学年': str(grade),
            '名前': f'{sei} {mei}',
            'ふりがな': f'{sei.lower()} {mei_k}',
            '正式名前': f'{sei} {mei}',
            '正式名前ふりがな': f'{sei.lower()} {mei_k}',
            '性別': sex,
            '生年月日': f'20{19-grade:02d}-{random.randint(4,12):02d}-{random.randint(1,28):02d}',
            # ... 住所・保護者等は適当に生成
        })
    return pd.DataFrame(rows)
```

---

## 付録

### A. 最優先アクションアイテム

| 優先度 | アクション | 担当 |
|--------|----------|------|
| ★★★ | C4th エクスポート Excel の実物サンプル入手 | RYUMA |
| ★★★ | Google Drive に配布フォルダを作成し、共有設定 | RYUMA |
| ★★☆ | IPAmj明朝のインストール確認（テスト用 PC） | RYUMA |
| ★★☆ | 花柄ボーダー画像素材の準備 | RYUMA |
| ★★☆ | 外字を含むテスト用ダミー名簿データ作成 | 開発者 |
| ★☆☆ | 保護者 2 以降のヘッダー形式確認 | RYUMA |

### B. 将来拡張

- 保護者 2 以降への対応
- 他帳票の追加（通知表、出席簿、保健関連等）
- テンプレートエディタ（教員が GUI でプレースホルダー配置）
- 複数校一括処理（教育委員会向け）
- C4th API 直接連携（将来の API 公開時）

---

> **文書終了** — 本文書は `CLAUDE.md` と合わせてプロジェクトルートに配置して使用する。
