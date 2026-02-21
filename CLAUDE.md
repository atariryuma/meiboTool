# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> 詳細仕様は `docs/SPEC.md` を参照。各章に仕様・実装コード・落とし穴がセットで記載されている。

## プロジェクト概要

那覇市立小中学校向けの **名簿帳票自動生成 Windows デスクトップアプリ**。
C4th（EDUCOM 校務支援システム）の Excel エクスポートを読み込み、名札・名列表・台帳・個票など 15 種の帳票を自動生成する。Python 3.11 + CustomTkinter + openpyxl。

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

# exe ビルド
pyinstaller build.spec
# → dist/名簿帳票ツール/ フォルダが出力される
```

## アーキテクチャ

### データフロー

```text
C4th Excel
  → core/importer.py  （ヘッダー行自動検出 → pandas DataFrame に読込）
  → core/mapper.py    （C4th カラム名 → 内部論理名に変換）
  → gui/frames/       （クラス選択パネルで学年・組を選択）
  → core/generator.py （テンプレートにデータを差込 → xlsx 出力）
```

### generator.py の 3 種類のジェネレーター

| クラス | 用途 | テンプレート |
| -------- | ------ | ------------ |
| `GridGenerator` | 1 ページに N 名をグリッド配置 | 名札・名列表・調べ表・横名簿 等 |
| `ListGenerator` | データ行を児童数分コピー展開 | 台帳 2 種 |
| `IndividualGenerator` | 1 名/シートで `copy_worksheet` 複製 | 個票 2 種（未テンプレート） |

`create_generator()` がファクトリー関数。`TEMPLATES` レジストリの `type` キーで分岐する。

### プレースホルダー形式

テンプレート Excel 内は `{{論理名}}` 形式。`fill_placeholders()` が正規表現で一括置換する。
特殊キー: `{{年度}}` `{{年度和暦}}` `{{学校名}}` `{{担任名}}` `{{住所}}` (住所 4 フィールド自動結合)。

### GUI 構成

```text
App (CTk) — 2カラムレイアウト
├── ヘッダーバー（設定ボタン）
├── 左パネル（CTkScrollableFrame）
│   ├── ImportFrame      — ファイル選択・読込件数・同期ステータス
│   ├── ClassSelectPanel — 学年・組の選択
│   ├── SelectFrame      — テンプレート選択・担任名・学校名（on_template_change コールバック）
│   └── OutputFrame      — 生成ボタン・進捗バー
└── 右パネル（_PreviewPanel — CTkTabview）
    ├── 「データ」タブ   — Treeview でデータ一覧表示
    └── 「プレビュー」タブ — PIL レンダリング画像で帳票プレビュー表示
```

### 自動更新

- **アプリ更新**: GitHub Releases API で最新バージョン確認 → zip ダウンロード → バッチファイルで更新
- **名簿データ同期**: LAN（共有フォルダ）/ Google Drive（AES-256-GCM 暗号化）の 2 モード対応
- **暗号化**: `core/crypto.py` — AES-256-GCM + PBKDF2-HMAC-SHA256 (600,000 回)、パスワードは Windows DPAPI で保護

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
3. **`copy_worksheet` は画像をコピーしない** — `_data()` でバイナリ取得 → `BytesIO` → `Image()` で再構築（`copy_sheet_with_images()` 参照）
4. **結合セルは左上セルにのみ書き込む** — 他セルへの書き込みは `AttributeError: 'MergedCell' is read-only`
5. **CustomTkinter の UI 操作はメインスレッドのみ** — バックグラウンドスレッドからは `self.after(0, callback)` で委譲する。ジェネレーター作成（Tkinter ウィジェット参照）もメインスレッドで行う
6. **C4th ヘッダーの全角スペース（U+3000）** — 「保護者1　続柄」等は全角スペース区切り。`normalize_header()` で統一する
7. **IPAmj明朝フォント** — 正式名前列に IVS 付き異体字を含むため必須。`font_helper.apply_font()` で全セルに適用
8. **IndividualGenerator のシート複製は fill 前に行う** — fill 後にコピーするとプレースホルダーが消失する

## テンプレート 15 種（カテゴリ別）

| カテゴリ | テンプレートファイル名 | タイプ | 状態 |
| -------- | ---------------------- | ------ | ---- |
| 名札・ラベル | ラベル_色付き.xlsx / サンプル_名札.xlsx / 名札_1年生用.xlsx | grid | ✅ |
| 名札・ラベル | ラベル_大2.xlsx / ラベル_小.xlsx / ラベル_特大.xlsx | grid | ✅ |
| 名簿・出欠表 | 掲示用名列表.xlsx / 調べ表.xlsx | grid | ✅ |
| 名簿・出欠表 | 横名簿.xlsx / 縦一週間.xlsx | grid | ✅ |
| 名簿・出欠表 | 男女一覧.xlsx | grid | ✅ sort_by=性別 |
| 台帳 | 修了台帳.xlsx / 卒業台帳.xlsx | list | ✅ |
| 個票 | 家庭調査票.xlsx / 学級編成用個票.xlsx | individual | ✅ |

## C4th データの重要な仕様

- **「組」「出席番号」は C4th に含まれない** — ファイル読込後にクラス選択パネルで教員が選択する（将来 C4th から出力予定）
- **特別支援学級** — C4th の「組」列に非数字の値（「なかよし」「ひまわり」等）が入る。`is_special_needs_class()` で判定。交流学級は初回インポート時にダイアログで割り当て → config に保存 → 以降自動統合
- **全カラム `dtype=str` で読み込む** — 生年月日等の型変換は後工程で行う
- **ヘッダー行自動検出** — ファイル先頭にメタ情報行がある可能性があるため、文字列セル 5 個以上の行を検索する

## config.json の主要キー

| キー | 説明 |
| ---- | ---- |
| `school_name` | 学校名（帳票タイトルに使用） |
| `school_type` | `"elementary"` / `"middle"` |
| `fiscal_year` | 年度（西暦・起動時に自動計算） |
| `template_dir` | テンプレートフォルダパス（`./テンプレート`） |
| `output_dir` | 出力フォルダパス（`./出力`） |
| `layout_dir` | レイアウトライブラリフォルダパス（`./レイアウト`） |
| `update.github_repo` | GitHub リポジトリ（`owner/repo` 形式） |
| `update.check_on_startup` | 起動時に更新チェックするか |
| `update.skip_version` | スキップ設定したバージョン |
| `data_source.mode` | `"manual"` / `"lan"` / `"gdrive"` |
| `data_source.lan_path` | LAN 共有フォルダの UNC パス |
| `data_source.gdrive_file_id` | Google Drive ファイル ID |
| `data_source.encryption_password` | DPAPI で保護された暗号化パスワード |
| `homeroom_teachers` | `{"1-1": "山田先生"}` 形式で担任名を保存 |
| `special_needs_assignments` | `{"1-なかよし-1": "1-1"}` 形式で特支児童の交流学級割り当てを保存 |
| `special_needs_placement` | `"appended"` (末尾追加) / `"integrated"` (出席番号順統合) |

## 現在地（セッション開始時に必ず確認）

> 詳細なフェーズ計画は `docs/ROADMAP.md` を参照。

### 実装済み ✅

| ファイル | 内容 | テスト |
| -------- | ---- | ------ |
| `utils/wareki.py` | 西暦→和暦変換 | ✅ 10 |
| `utils/address.py` | 住所4フィールド結合 | ✅ 5 |
| `utils/date_fmt.py` | 日付フォーマット（YY/MM/DD・Excelシリアル値対応） | ✅ 19 |
| `utils/font_helper.py` | IPAmj明朝フォント適用 | ✅ 11 |
| `core/config.py` | config.json 読み書き・deep_merge・パス解決 | ✅ 13 |
| `core/mapper.py` | C4th カラム名マッピング + resolve_name_fields | ✅ 12 |
| `core/importer.py` | ヘッダー自動検出付き Excel 読込 | ✅ 11 |
| `core/generator.py` | Grid/List/Individual ジェネレーター + 性別ソート + 画像警告抑制 | ✅ 52 |
| `core/crypto.py` | AES-256-GCM 暗号化/復号 + DPAPI パスワード保護 | ✅ 17 |
| `core/data_sync.py` | 名簿データ自動同期（LAN/GDrive/手動） | ✅ 15 |
| `core/updater.py` | GitHub Releases ベースのアプリ更新 | ✅ 20 |
| `templates/template_registry.py` | テンプレート 15 種メタデータ + カテゴリ別表示 | ✅ 30 |
| `templates/template_scanner.py` | テンプレート自動検出（.xlsx スキャン → メタデータ推定） | ✅ 13 |
| `templates/generators/gen_meireihyo.py` | 掲示用名列表テンプレート生成 | ✅ 12 |
| `templates/generators/gen_nafuda.py` | 名札3種テンプレート生成 | ✅ 16 |
| `templates/generators/gen_daicho.py` | 台帳2種テンプレート生成 | ✅ 14 |
| `templates/generators/gen_shirabehyo.py` | 調べ表テンプレート生成 | ✅ 12 |
| `templates/generators/gen_katei_chousahyo.py` | 家庭調査票テンプレート生成 | ✅ 19 |
| `templates/generators/gen_gakkyuu_kojihyo.py` | 学級編成用個票テンプレート生成 | ✅ 19 |
| `templates/generators/gen_from_legacy.py` | レガシーテンプレート変換 | — |
| `templates/generators/generate_all.py` | 全テンプレート一括生成 | — |
| `core/special_needs.py` | 特別支援学級判定・検出・統合・交流学級割り当てロジック | ✅ 36 |
| `gui/app.py` | メインウィンドウ（2カラム + CTkTabview プレビュー + 特支自動統合） | ✅ 35 |
| `gui/preview_renderer.py` | openpyxl Worksheet → PIL Image レンダラー（IPAmj明朝・画像・垂直揃え対応） | ✅ 22 |
| `gui/frames/import_frame.py` | ファイル選択・同期ステータス表示 | — |
| `gui/frames/class_select_panel.py` | 学年・組選択 + 特別支援学級表示 | — |
| `gui/frames/select_frame.py` | テンプレート選択（カテゴリ別）・担任名・学校名・特支配置設定 | — |
| `gui/frames/output_frame.py` | 生成ボタン・進捗バー | — |
| `gui/dialogs/mapping_dialog.py` | カラムマッピング手動調整ダイアログ | ✅ 13 |
| `gui/dialogs/settings_dialog.py` | 管理者設定（同期モード設定） | — |
| `gui/dialogs/exchange_class_dialog.py` | 交流学級割り当てダイアログ | — |
| `gui/dialogs/update_dialog.py` | 更新確認ダイアログ | — |
| `core/lay_parser.py` | .lay バイナリパーサー + FontInfo bold/italic + `new_label/new_field/new_line` | ✅ 29 |
| `core/lay_serializer.py` | LayFile ↔ JSON 保存/読込（bold/italic ラウンドトリップ） | ✅ 22 |
| `core/lay_renderer.py` | Canvas/PILBackend 描画 + 縦書き/複数行 + `fill_layout()` + `render_layout_to_image()` | ✅ 50 |
| `core/win_printer.py` | Windows GDI 直接印刷（縦書き・複数行・bold/italic 対応） | ✅ 13 |
| `core/layout_registry.py` | レイアウトライブラリ管理（scan/import/delete/rename） | ✅ 15 |
| `gui/editor/editor_window.py` | レイアウトエディター メインウィンドウ（Canvas + プロパティ + Undo/Redo） | — |
| `gui/editor/layout_canvas.py` | インタラクティブ Canvas（選択・移動・リサイズ） | — |
| `gui/editor/properties_panel.py` | オブジェクトプロパティ編集パネル（bold/italic 保持） | — |
| `gui/editor/toolbar.py` | ツールバー（ファイル操作・ズーム・データ差込・ライブラリ保存） | — |
| `gui/editor/data_fill_dialog.py` | C4th データ差込ダイアログ | — |
| `gui/editor/print_dialog.py` | プリンター選択・印刷実行ダイアログ | — |
| `gui/editor/print_preview_dialog.py` | 印刷プレビュー（PILBackend レンダリング + ページ送り） | — |
| `gui/editor/layout_manager_dialog.py` | レイアウトライブラリ管理ダイアログ（Treeview 一覧） | — |
| `gui/editor/object_list.py` | オブジェクト一覧パネル（TreeView + Canvas 連動） | — |
| `templates/generators/gen_from_lay.py` | .lay → Excel テンプレート変換 | ✅ 27 |
| `.github/workflows/build-release.yml` | CI/CD（テスト→ビルド→Release） | — |
| `tests/conftest.py` | 共通フィクスチャ | — |

### 未実装 ❌

1. 別 PC（Python なし環境）での exe 起動確認

### 開発環境の状態

- テスト: **646 件全パス**（`venv/Scripts/python.exe -m pytest`）
- リント: ruff クリーン（`venv/Scripts/python.exe -m ruff check meibo_tool/`）
- Git: master ブランチ

## 詳細仕様の参照先（docs/SPEC.md）

| 内容 | 章 |
| ---- | -- |
| C4th 全 50 カラム定義・EXACT_MAP | 第2章 |
| GUI 各パネルの詳細仕様・実装コード | 第3章 |
| テンプレート 15 種の列幅・行高・差込フィールド | 第4章 |
| generator.py・fill_placeholders・印刷設定 | 第5章 |
| 自動更新（GitHub Releases + 名簿データ同期） | 第6章 |
| build.spec・テストケース一覧・ダミーデータ生成 | 第7章 |
