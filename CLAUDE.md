# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> 詳細仕様は `docs/SPEC.md` を参照。各章に仕様・実装コード・落とし穴がセットで記載されている。

## プロジェクト概要

那覇市立小中学校向けの **名簿帳票自動生成 Windows デスクトップアプリ**。
C4th（EDUCOM 校務支援システム）の Excel エクスポートを読み込み、名札・名列表・台帳・個票など 9 種の帳票を自動生成する。Python 3.11 + CustomTkinter + openpyxl。

## コマンド

```bash
# 仮想環境セットアップ（初回のみ）
python -m venv venv
venv/Scripts/activate      # Windows Git Bash / Mac: source venv/bin/activate
pip install -r requirements.txt

# アプリ起動（プロジェクトルートから実行）
cd meibo_tool && python main.py

# テスト全実行（プロジェクトルートから。pyproject.toml で testpaths 設定済み）
venv/Scripts/python.exe -m pytest

# 単一テストファイル実行
venv/Scripts/python.exe -m pytest meibo_tool/tests/test_wareki.py -v

# リント（ゼロエラーになってからコミット）
venv/Scripts/python.exe -m ruff check meibo_tool/
venv/Scripts/python.exe -m ruff check meibo_tool/ --fix   # 自動修正

# テンプレート Excel 生成（初回および更新時）
cd meibo_tool && python -m templates.generators.generate_all

# exe ビルド（build.spec が存在する場合）
pyinstaller build.spec
# → dist/名簿帳票ツール/ フォルダが出力される
```

## アーキテクチャ

### データフロー

```text
C4th Excel
  → core/importer.py  （ヘッダー行自動検出 → pandas DataFrame に読込）
  → core/mapper.py    （C4th カラム名 → 内部論理名に変換）
  → gui/frames/       （必須入力パネルで「組」「出席番号」を追加）
  → core/generator.py （テンプレートにデータを差込 → xlsx 出力）
```

### generator.py の 3 種類のジェネレーター

| クラス | 用途 | テンプレート |
| -------- | ------ | ------------ |
| `GridGenerator` | 1 ページに N 名をグリッド配置 | 名札 3 種 |
| `ListGenerator` | データ行を児童数分コピー展開 | 名列表・台帳 4 種 |
| `IndividualGenerator` | 1 名/シートで `copy_worksheet` 複製 | 個票 2 種 |

`create_generator()` がファクトリー関数。`TEMPLATES` レジストリの `type` キーで分岐する。

### プレースホルダー形式

テンプレート Excel 内は `{{論理名}}` 形式。`fill_placeholders()` が正規表現で一括置換する。
特殊キー: `{{年度}}` `{{年度和暦}}` `{{学校名}}` `{{担任名}}` `{{住所}}` (住所 4 フィールド自動結合)。

### GUI 状態管理

`App.mandatory_ok` フラグが `False` の間、テンプレート選択・生成ボタンは `state='disabled'`。
`MandatoryInputPanel` で「組」選択＋「自動連番」クリック後に「確定して進む」を押すと `True` になる。

## コーディング規約

- **リンター**: ruff。コード変更後は必ず `ruff check meibo_tool/` でゼロエラーを確認してからコミット
- **import 順序**: 標準ライブラリ → サードパーティ → 社内モジュール（ruff の I001 で自動チェック）
- **未使用ループ変数**: `_name` のようにアンダースコアプレフィックスを付ける
- **型注釈**: 関数シグネチャには型注釈を付ける（`def func(x: str) -> int:`）
- **テスト**: 新機能は同時に `tests/test_<モジュール名>.py` に追加
- **sys.path ハック禁止**: テストファイルに `sys.path.insert` を書かない（`pyproject.toml` の `pythonpath` で解決済み）
- **共通フィクスチャ**: `meibo_tool/tests/conftest.py` に追加（`dummy_df`・`default_options`・`default_config`）

## 絶対守るべきルール

1. **PyInstaller は `--onedir` のみ** — CustomTkinter の .json/.otf が `--onefile` では見つからずクラッシュする
2. **openpyxl スタイルはイミュータブル** — `cell.font.bold = True` は AttributeError。必ず `cell.font = Font(...)` で新インスタンスを作る
3. **`copy_worksheet` は画像をコピーしない** — `source_ws._images` を手動でループして `add_image()` で再挿入する
4. **結合セルは左上セルにのみ書き込む** — 他セルへの書き込みは `AttributeError: 'MergedCell' is read-only`
5. **CustomTkinter の UI 操作はメインスレッドのみ** — バックグラウンドスレッドからは `self.after(0, callback)` で委譲する
6. **C4th ヘッダーの全角スペース（U+3000）** — 「保護者1　続柄」等は全角スペース区切り。`normalize_header()` で統一する
7. **IPAmj明朝フォント** — 正式名前列に IVS 付き異体字を含むため必須。`font_helper.apply_font()` で全セルに適用

## テンプレート 9 種

| テンプレートファイル名 | タイプ | use_formal_name |
| ---------------------- | ------ | --------------- |
| 名札_通常.xlsx / 名札_装飾あり.xlsx / 名札_1年生用.xlsx | grid | false |
| 掲示用名列表.xlsx / 調べ表.xlsx | list | false |
| 修了台帳.xlsx / 卒業台帳.xlsx | list | true |
| 家庭調査票.xlsx / 学級編成用個票.xlsx | individual | true |

`use_formal_name=true` のテンプレートでは「正式名前」列（IVS 付き）を使用。空の場合は「名前」にフォールバックする（`resolve_name_fields()` 参照）。

## C4th データの重要な仕様

- **「組」「出席番号」は C4th に含まれない** — ファイル読込後に必須入力パネルで教員が入力する
- **全カラム `dtype=str` で読み込む** — 生年月日等の型変換は後工程で行う
- **ヘッダー行自動検出** — ファイル先頭にメタ情報行がある可能性があるため、文字列セル 5 個以上の行を検索する

## config.json の主要キー

| キー | 説明 |
| ---- | ---- |
| `school_name` | 学校名（帳票タイトルに使用） |
| `school_type` | `"elementary"` / `"middle"` |
| `fiscal_year` | 年度（西暦） |
| `template_dir` | テンプレートフォルダパス（`./テンプレート`） |
| `output_dir` | 出力フォルダパス（`./出力`） |
| `update.version_file_id` | Google Drive の version.json のファイル ID |
| `homeroom_teachers` | `{"1-1": "山田先生"}` 形式で担任名を保存 |

## Google Drive 自動更新

- 起動時にバックグラウンドスレッドで `updater.check_for_updates()` を実行
- `version.json` を API キー不要で取得（「リンクを知っている全員が閲覧可」設定前提）
- exe 自己更新はバッチファイル経由（実行中の exe はロックされるため）
- バッチファイルには `chcp 65001` 必須（日本語パスの文字化け対策）
- ネットワーク失敗時はサイレントにスキップし、通常起動する

## 現在地（セッション開始時に必ず確認）

> 詳細なフェーズ計画は `docs/ROADMAP.md` を参照。

### 実装済み ✅

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `utils/wareki.py` | 西暦→和暦変換 | ✅ 10ケース |
| `utils/address.py` | 住所4フィールド結合 | ✅ 5ケース |
| `utils/font_helper.py` | IPAmj明朝フォント適用 | — |
| `core/config.py` | config.json 読み書き | — |
| `core/mapper.py` | C4th カラム名マッピング（全50列） | ✅ 6ケース |
| `core/importer.py` | ヘッダー自動検出付き Excel 読込 | — |
| `core/generator.py` | Grid/List/Individual ジェネレーター骨格 | — |
| `templates/template_registry.py` | テンプレート 9 種メタデータ | — |
| `tests/conftest.py` | 共通フィクスチャ（dummy_df 等） | — |

### 未実装 ❌（次に着手する順）

1. **`templates/generators/gen_meireihyo.py`** ← **次のタスク**
   - 掲示用名列表テンプレート Excel を openpyxl でプログラム生成
   - `テンプレート/掲示用名列表.xlsx` が出力される
2. `tests/test_generator.py` — ListGenerator × 名列表の統合テスト
3. `gui/app.py` + `gui/frames/` — CustomTkinter GUI
4. 名札テンプレート 3 種 (`gen_nafuda_*.py`)
5. 台帳テンプレート 2 種 + 調べ表 (`gen_*_daicho.py`, `gen_shirabehyo.py`)
6. 個票テンプレート 2 種 (`gen_katei_*.py`, `gen_gakkyuu_*.py`)
7. `core/updater.py` — Google Drive 自動更新
8. `build.spec` + PyInstaller ビルド

### 開発環境の状態

- テスト: 21件 全パス（`venv/Scripts/python.exe -m pytest`）
- リント: ruff クリーン（`venv/Scripts/python.exe -m ruff check meibo_tool/`）
- Git: 3コミット済み（master ブランチ）

## 実装順序（依存関係順）

1. `utils/wareki.py` → `utils/address.py` → `utils/font_helper.py`
2. `core/config.py` → `core/mapper.py` → `core/importer.py`
3. `templates/template_registry.py` → テンプレート生成スクリプト（`gen_meireihyo.py` を最初に）
4. `core/generator.py`（ListGenerator から） → `gui/` → 統合テスト
5. `core/updater.py` → `build.spec` + ビルド

## 詳細仕様の参照先（docs/SPEC.md）

| 内容 | 章 |
| ---- | -- |
| C4th 全 50 カラム定義・EXACT_MAP | 第2章 |
| GUI 各パネルの詳細仕様・実装コード | 第3章 |
| テンプレート 9 種の列幅・行高・差込フィールド | 第4章 |
| generator.py・fill_placeholders・印刷設定 | 第5章 |
| Google Drive 更新・exe 自己更新バッチ | 第6章 |
| build.spec・テストケース一覧・ダミーデータ生成 | 第7章 |
